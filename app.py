from __future__ import annotations

import argparse
import json
import os
import queue
import shutil
import subprocess
import tempfile
import threading
import traceback
import webbrowser
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter.scrolledtext import ScrolledText

from PIL import Image, ImageTk

from about_profile import ABOUT_PROFILE_PATH, load_about_profile, resolve_profile_image
from app_state import APP_NAME, APP_STATE_PATH, AppStateStore
from engagement_core import (
    ensure_install_date,
    iso_now,
    parse_datetime,
    should_show_first_launch_splash,
    should_show_login_popup,
    summarize_login_popup_state,
)
from engagement_ui import FirstLaunchSplashWindow, LoginReminderToast
from automation_core import (
    WatchFolderConfig,
    bundle_paths_as_zip,
    discover_watch_candidates,
    export_presets_to_json,
    import_presets_from_json,
    move_files_to_archive,
    normalize_watch_config,
    write_run_report,
)
from build_support import export_diagnostics_report, export_state_snapshot, import_state_snapshot, export_text_file
from converter_core import (
    BatchConfig,
    PdfToolConfig,
    ENGINE_AUTO,
    ENGINE_HELP,
    ENGINE_LIBREOFFICE,
    ENGINE_ORDER,
    ENGINE_PURE_PYTHON,
    MODE_ANY_TO_PDF,
    MODE_DOCS_TO_PDF,
    MODE_HELP,
    MODE_HTML_TO_DOCX,
    MODE_HTML_TO_MD,
    MODE_HTML_TO_PDF,
    MODE_IMAGES_TO_PDF,
    MODE_MD_TO_DOCX,
    MODE_MD_TO_HTML,
    MODE_MD_TO_PDF,
    MODE_MERGE_PDFS,
    MODE_PDF_TO_DOCX,
    MODE_PDF_TO_HTML,
    MODE_PDF_TO_IMAGES,
    MODE_PDF_TO_PPTX,
    MODE_PDF_TO_XLSX,
    MODE_PRESENTATIONS_TO_IMAGES,
    MODE_PRESENTATIONS_TO_PDF,
    MODE_SHEETS_TO_PDF,
    MODE_TEXT_TO_PDF,
    PDF_TOOL_COMPRESS,
    PDF_TOOL_EDIT_METADATA,
    PDF_TOOL_EXTRACT_PAGES,
    PDF_TOOL_HELP,
    PDF_TOOL_LOCK,
    PDF_TOOL_IMAGE_OVERLAY,
    PDF_TOOL_REDACT_AREA,
    PDF_TOOL_EDIT_TEXT,
    PDF_TOOL_MERGE,
    PDF_TOOL_ORDER,
    PDF_TOOL_REDACT_TEXT,
    PDF_TOOL_REMOVE_PAGES,
    PDF_TOOL_REORDER_PAGES,
    PDF_TOOL_SIGN_VISIBLE,
    PDF_TOOL_UNLOCK,
    PDF_TOOL_SPLIT_EVERY_N,
    PDF_TOOL_SPLIT_RANGES,
    PDF_TOOL_TEXT_OVERLAY,
    PDF_TOOL_WATERMARK_IMAGE,
    PDF_TOOL_WATERMARK_TEXT,
    collect_files_from_folder,
    default_merged_name,
    dependency_status,
    build_conversion_route_preview,
    filetype_patterns_for_mode,
    outputs_pdf,
    process_batch,
    process_pdf_tool,
    supported_extensions_for_mode,
)
from page_organizer import PageOrganizerPanel
from mail_core import SMTPSettings, build_eml_draft, create_mailto_url, open_mailto_draft, send_email, test_smtp_connection
from link_ingest import cache_root_from_setting, clear_cache_dir, download_many_urls, extract_urls
from ocr_core import OcrConfig, OcrError, detect_tesseract_status, extract_text_with_ocr, image_to_searchable_pdf, pdf_to_searchable_pdf
from workflow_support import directory_stats, format_bytes, prune_directory, summarize_directory
from workflow_ui import CommandPaletteWindow, QuickAction, Tooltip
from ui_theme import (
    ThemePalette,
    apply_menu_theme,
    apply_text_widget_theme,
    apply_treeview_tag_colors,
    apply_ttk_theme,
    resolve_palette,
)

APP_VERSION = "1.5.0 Patch 15"
MODE_ORDER = [
    MODE_ANY_TO_PDF,
    MODE_IMAGES_TO_PDF,
    MODE_PDF_TO_IMAGES,
    MODE_DOCS_TO_PDF,
    MODE_PRESENTATIONS_TO_PDF,
    MODE_PRESENTATIONS_TO_IMAGES,
    MODE_PDF_TO_DOCX,
    MODE_PDF_TO_PPTX,
    MODE_PDF_TO_HTML,
    MODE_SHEETS_TO_PDF,
    MODE_PDF_TO_XLSX,
    MODE_TEXT_TO_PDF,
    MODE_MD_TO_PDF,
    MODE_MD_TO_DOCX,
    MODE_MD_TO_HTML,
    MODE_HTML_TO_PDF,
    MODE_HTML_TO_DOCX,
    MODE_HTML_TO_MD,
    MODE_MERGE_PDFS,
]
NAV_PAGES = [
    ("home", "Home"),
    ("convert", "Convert"),
    ("pdf_tools", "PDF Tools"),
    ("ocr", "OCR"),
    ("organizer", "Organizer"),
    ("automation", "Automation"),
    ("history", "History"),
    ("settings", "Settings"),
    ("about", "About"),
]
PDF_TOOL_POSITIONS = ["center", "top-left", "top-right", "bottom-left", "bottom-right"]


class MarkdownNotesWindow(tk.Toplevel):
    def __init__(self, master: tk.Misc, notes_path: Path, palette: ThemePalette) -> None:
        super().__init__(master)
        self.notes_path = notes_path
        self.palette = palette
        self.title(f"{APP_NAME} - Footer Notes")
        self.geometry("760x620")
        self.minsize(620, 420)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ttk.Frame(self, style="Surface.TFrame", padding=18)
        header.grid(row=0, column=0, sticky="nsew")
        header.grid_columnconfigure(0, weight=1)

        ttk.Label(header, text="Footer notes from Markdown", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Edit footer_notes.md any time. Reload to reflect your latest content.",
            style="CardBody.TLabel",
            wraplength=560,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        actions = ttk.Frame(header, style="Surface.TFrame")
        actions.grid(row=0, column=1, rowspan=2, sticky="e")
        ttk.Button(actions, text="Reload", command=self.reload_notes).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(actions, text="Open Markdown File", command=self.open_notes_file).grid(row=0, column=1)

        self.viewer = ScrolledText(self, wrap=tk.WORD, relief="flat", borderwidth=1, padx=16, pady=14)
        self.viewer.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 18))
        self.viewer.configure(state="disabled")

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.apply_theme(palette)
        self.reload_notes()

    def apply_theme(self, palette: ThemePalette) -> None:
        self.palette = palette
        apply_ttk_theme(self, palette)
        apply_text_widget_theme(self.viewer, palette)
        self.viewer.configure(font=("Consolas", 10) if os.name == "nt" else ("TkFixedFont", 10))
        self.viewer.tag_configure("h1", font=("Segoe UI", 16, "bold"), spacing1=14, spacing3=10)
        self.viewer.tag_configure("h2", font=("Segoe UI", 13, "bold"), spacing1=12, spacing3=8)
        self.viewer.tag_configure("body", lmargin1=4, lmargin2=4, spacing3=3)
        self.viewer.tag_configure("bullet", lmargin1=18, lmargin2=32, spacing3=2)
        self.viewer.tag_configure("quote", lmargin1=18, lmargin2=32, foreground=palette.text_muted)
        self.viewer.tag_configure("code", background=palette.surface_alt, font=("Consolas", 10) if os.name == "nt" else ("TkFixedFont", 10))
        self.viewer.tag_configure("muted", foreground=palette.text_muted)

    def reload_notes(self) -> None:
        if not self.notes_path.exists():
            self.notes_path.write_text(
                "# Footer Notes\n\nUpdate this file with anything you want to show in the footer notes window.\n",
                encoding="utf-8",
            )
        content = self.notes_path.read_text(encoding="utf-8")
        self._render_markdown(content)

    def _render_markdown(self, content: str) -> None:
        self.viewer.configure(state="normal")
        self.viewer.delete("1.0", tk.END)

        in_code = False
        for raw_line in content.splitlines():
            line = raw_line.rstrip("\n")
            stripped = line.strip()

            if stripped.startswith("```"):
                in_code = not in_code
                continue

            if in_code:
                self.viewer.insert(tk.END, line + "\n", ("code",))
                continue

            if stripped.startswith("# "):
                self.viewer.insert(tk.END, stripped[2:] + "\n", ("h1",))
            elif stripped.startswith("## "):
                self.viewer.insert(tk.END, stripped[3:] + "\n", ("h2",))
            elif stripped.startswith(">"):
                self.viewer.insert(tk.END, stripped.lstrip("> ") + "\n", ("quote",))
            elif stripped.startswith(("- ", "* ")):
                self.viewer.insert(tk.END, "• " + stripped[2:] + "\n", ("bullet",))
            elif not stripped:
                self.viewer.insert(tk.END, "\n", ("body",))
            else:
                self.viewer.insert(tk.END, line + "\n", ("body",))

        self.viewer.configure(state="disabled")
        self.viewer.see("1.0")

    def open_notes_file(self) -> None:
        open_path(self.notes_path)


class SMTPSettingsWindow(tk.Toplevel):
    def __init__(self, app: "GokulOmniConvertLiteApp", palette: ThemePalette) -> None:
        super().__init__(app)
        self.app = app
        self.palette = palette
        self.title(f"{APP_NAME} - SMTP Delivery")
        self.geometry("760x640")
        self.minsize(640, 520)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ttk.Frame(self, style="Card.TFrame", padding=18)
        header.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 10))
        header.grid_columnconfigure(0, weight=1)
        ttk.Label(header, text="Direct email delivery", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text=(
                "Configure SMTP once, test the connection, then send the latest outputs directly from the app. "
                "If you enable password saving, it is stored locally in your app state file on this machine."
            ),
            style="CardBody.TLabel",
            wraplength=620,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        body = ttk.Frame(self, style="Surface.TFrame", padding=18)
        body.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 10))
        body.grid_columnconfigure(1, weight=1)
        body.grid_columnconfigure(3, weight=1)

        row = 0
        ttk.Label(body, text="SMTP host:", style="Surface.TLabel").grid(row=row, column=0, sticky="w")
        ttk.Entry(body, textvariable=self.app.smtp_host_var).grid(row=row, column=1, sticky="ew", padx=(8, 16))
        ttk.Label(body, text="Port:", style="Surface.TLabel").grid(row=row, column=2, sticky="w")
        ttk.Entry(body, textvariable=self.app.smtp_port_var, width=10).grid(row=row, column=3, sticky="w", padx=(8, 0))

        row += 1
        ttk.Label(body, text="Username:", style="Surface.TLabel").grid(row=row, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(body, textvariable=self.app.smtp_username_var).grid(row=row, column=1, sticky="ew", padx=(8, 16), pady=(10, 0))
        ttk.Label(body, text="Password:", style="Surface.TLabel").grid(row=row, column=2, sticky="w", pady=(10, 0))
        ttk.Entry(body, textvariable=self.app.smtp_password_var, show="*").grid(row=row, column=3, sticky="ew", padx=(8, 0), pady=(10, 0))

        row += 1
        ttk.Label(body, text="Sender:", style="Surface.TLabel").grid(row=row, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(body, textvariable=self.app.smtp_sender_var).grid(row=row, column=1, sticky="ew", padx=(8, 16), pady=(10, 0))
        ttk.Label(body, text="Default recipient:", style="Surface.TLabel").grid(row=row, column=2, sticky="w", pady=(10, 0))
        ttk.Entry(body, textvariable=self.app.smtp_default_to_var).grid(row=row, column=3, sticky="ew", padx=(8, 0), pady=(10, 0))

        row += 1
        options = ttk.Frame(body, style="Surface.TFrame")
        options.grid(row=row, column=0, columnspan=4, sticky="ew", pady=(14, 0))
        ttk.Checkbutton(options, text="Use SSL", variable=self.app.smtp_use_ssl_var).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(options, text="Use STARTTLS", variable=self.app.smtp_use_starttls_var).grid(row=0, column=1, sticky="w", padx=(14, 0))
        ttk.Checkbutton(options, text="Save SMTP password locally", variable=self.app.smtp_save_password_var).grid(row=0, column=2, sticky="w", padx=(14, 0))

        row += 1
        ttk.Label(body, textvariable=self.app.smtp_status_var, style="CardBody.TLabel", wraplength=620, justify="left").grid(
            row=row, column=0, columnspan=4, sticky="w", pady=(12, 0)
        )

        actions = ttk.Frame(self, style="Surface.TFrame", padding=(18, 0, 18, 18))
        actions.grid(row=2, column=0, sticky="ew")
        ttk.Button(actions, text="Test connection", command=self.app._test_smtp_settings).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(actions, text="Send latest outputs", style="Primary.TButton", command=self.app._send_last_outputs_via_smtp).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(actions, text="Open mail draft", command=self.app._open_mail_draft_for_last_outputs).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(actions, text="Save settings", command=self.app._save_smtp_settings).grid(row=0, column=3, padx=(0, 8))
        ttk.Button(actions, text="Clear password", command=self._clear_password).grid(row=0, column=4)

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.apply_theme(palette)

    def _clear_password(self) -> None:
        self.app.smtp_password_var.set("")
        self.app.smtp_save_password_var.set(False)
        self.app.smtp_status_var.set("SMTP password cleared from the current session.")

    def apply_theme(self, palette: ThemePalette) -> None:
        self.palette = palette
        apply_ttk_theme(self, palette)


class BuildCenterWindow(tk.Toplevel):
    def __init__(self, app: "GokulOmniConvertLiteApp", palette: ThemePalette) -> None:
        super().__init__(app)
        self.app = app
        self.palette = palette
        self.title(f"{APP_NAME} - Build Center")
        self.geometry("760x620")
        self.minsize(620, 500)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ttk.Frame(self, style="Card.TFrame", padding=18)
        header.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 10))
        header.grid_columnconfigure(0, weight=1)
        ttk.Label(header, text="Installer and release prep", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text=(
                "Build Center groups diagnostics, state snapshots, installer notes, and packaging shortcuts so release work stays organized."
            ),
            style="CardBody.TLabel",
            wraplength=620,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        body = ttk.Frame(self, style="Surface.TFrame", padding=18)
        body.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 10))
        body.grid_columnconfigure(0, weight=1)

        self.summary_label = ttk.Label(body, style="CardBody.TLabel", wraplength=620, justify="left")
        self.summary_label.grid(row=0, column=0, sticky="w")

        actions = ttk.Frame(body, style="Surface.TFrame")
        actions.grid(row=1, column=0, sticky="ew", pady=(16, 0))
        for column in range(2):
            actions.grid_columnconfigure(column, weight=1)

        ttk.Button(actions, text="Export diagnostics JSON", command=self.app._export_diagnostics_report_action).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(actions, text="Export settings snapshot", command=self.app._export_state_snapshot_action).grid(row=0, column=1, sticky="ew")
        ttk.Button(actions, text="Import settings snapshot", command=self.app._import_state_snapshot_action).grid(row=1, column=0, sticky="ew", padx=(0, 8), pady=(8, 0))
        ttk.Button(actions, text="Open installer folder", command=self.app._open_installer_folder).grid(row=1, column=1, sticky="ew", pady=(8, 0))
        ttk.Button(actions, text="Open build notes", command=lambda: open_path(self.app.build_notes_path)).grid(row=2, column=0, sticky="ew", padx=(0, 8), pady=(8, 0))
        ttk.Button(actions, text="Open SMTP settings", command=self.app._open_smtp_window).grid(row=2, column=1, sticky="ew", pady=(8, 0))
        ttk.Button(actions, text="Open About editor", command=self.app._open_about_editor_window).grid(row=3, column=0, sticky="ew", padx=(0, 8), pady=(8, 0))
        ttk.Button(actions, text="Refresh summary", command=self.refresh_summary).grid(row=3, column=1, sticky="ew", pady=(8, 0))

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.apply_theme(palette)
        self.refresh_summary()

    def refresh_summary(self) -> None:
        installer_root = self.app.build_notes_path.parent
        asset_count = sum(1 for path in installer_root.rglob("*") if path.is_file()) if installer_root.exists() else 0
        summary = [
            f"Version: {APP_VERSION}",
            f"Installer assets found: {asset_count}",
            f"App state file: {APP_STATE_PATH}",
            f"Current output folder: {self.app.output_dir_var.get().strip()}",
            f"Latest outputs tracked: {len(self.app.last_outputs)}",
            f"Dependency summary: {self.app.dependency_var.get().strip()}",
        ]
        self.summary_label.configure(text="\n\n".join(summary))

    def apply_theme(self, palette: ThemePalette) -> None:
        self.palette = palette
        apply_ttk_theme(self, palette)



class AboutProfileEditorWindow(tk.Toplevel):
    def __init__(self, app: "GokulOmniConvertLiteApp", palette: ThemePalette) -> None:
        super().__init__(app)
        self.app = app
        self.palette = palette
        self.title(f"{APP_NAME} - Edit About Profile")
        self.geometry("900x780")
        self.minsize(760, 620)

        self.name_var = tk.StringVar()
        self.title_var = tk.StringVar()
        self.subtitle_var = tk.StringVar()
        self.company_var = tk.StringVar()
        self.project_var = tk.StringVar()
        self.email_var = tk.StringVar()
        self.handle_var = tk.StringVar()
        self.image_path_var = tk.StringVar()
        self.feedback_url_var = tk.StringVar()
        self.contribute_url_var = tk.StringVar()
        self.link_label_vars = [tk.StringVar() for _ in range(5)]
        self.link_url_vars = [tk.StringVar() for _ in range(5)]

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ttk.Frame(self, style="Card.TFrame", padding=18)
        header.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 10))
        header.grid_columnconfigure(0, weight=1)
        ttk.Label(header, text="Edit About profile", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text=(
                "Update your name, project, company, feedback links, image path, bio, and social links here. "
                "Saving writes back to about_profile.json and refreshes the About page immediately."
            ),
            style="CardBody.TLabel",
            wraplength=760,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        body = ttk.Frame(self, style="Surface.TFrame", padding=18)
        body.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 10))
        body.grid_columnconfigure(1, weight=1)
        body.grid_columnconfigure(3, weight=1)

        row = 0
        ttk.Label(body, text="Name:", style="Surface.TLabel").grid(row=row, column=0, sticky="w")
        ttk.Entry(body, textvariable=self.name_var).grid(row=row, column=1, sticky="ew", padx=(8, 16))
        ttk.Label(body, text="Title:", style="Surface.TLabel").grid(row=row, column=2, sticky="w")
        ttk.Entry(body, textvariable=self.title_var).grid(row=row, column=3, sticky="ew", padx=(8, 0))

        row += 1
        ttk.Label(body, text="Subtitle:", style="Surface.TLabel").grid(row=row, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(body, textvariable=self.subtitle_var).grid(row=row, column=1, sticky="ew", padx=(8, 16), pady=(10, 0))
        ttk.Label(body, text="Handle:", style="Surface.TLabel").grid(row=row, column=2, sticky="w", pady=(10, 0))
        ttk.Entry(body, textvariable=self.handle_var).grid(row=row, column=3, sticky="ew", padx=(8, 0), pady=(10, 0))

        row += 1
        ttk.Label(body, text="Company:", style="Surface.TLabel").grid(row=row, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(body, textvariable=self.company_var).grid(row=row, column=1, sticky="ew", padx=(8, 16), pady=(10, 0))
        ttk.Label(body, text="Project:", style="Surface.TLabel").grid(row=row, column=2, sticky="w", pady=(10, 0))
        ttk.Entry(body, textvariable=self.project_var).grid(row=row, column=3, sticky="ew", padx=(8, 0), pady=(10, 0))

        row += 1
        ttk.Label(body, text="Email:", style="Surface.TLabel").grid(row=row, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(body, textvariable=self.email_var).grid(row=row, column=1, sticky="ew", padx=(8, 16), pady=(10, 0))
        ttk.Label(body, text="Image path:", style="Surface.TLabel").grid(row=row, column=2, sticky="w", pady=(10, 0))
        image_frame = ttk.Frame(body, style="Surface.TFrame")
        image_frame.grid(row=row, column=3, sticky="ew", padx=(8, 0), pady=(10, 0))
        image_frame.grid_columnconfigure(0, weight=1)
        ttk.Entry(image_frame, textvariable=self.image_path_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(image_frame, text="Browse", command=self._browse_image).grid(row=0, column=1, padx=(8, 0))

        row += 1
        ttk.Label(body, text="Feedback URL:", style="Surface.TLabel").grid(row=row, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(body, textvariable=self.feedback_url_var).grid(row=row, column=1, sticky="ew", padx=(8, 16), pady=(10, 0))
        ttk.Label(body, text="Contribute URL:", style="Surface.TLabel").grid(row=row, column=2, sticky="w", pady=(10, 0))
        ttk.Entry(body, textvariable=self.contribute_url_var).grid(row=row, column=3, sticky="ew", padx=(8, 0), pady=(10, 0))

        row += 1
        ttk.Label(body, text="Bio:", style="Surface.TLabel").grid(row=row, column=0, sticky="nw", pady=(12, 0))
        self.bio_text = ScrolledText(body, wrap=tk.WORD, height=7, relief="flat", borderwidth=1, padx=8, pady=8)
        self.bio_text.grid(row=row, column=1, columnspan=3, sticky="nsew", padx=(8, 0), pady=(12, 0))

        row += 1
        ttk.Separator(body, orient="horizontal").grid(row=row, column=0, columnspan=4, sticky="ew", pady=14)

        row += 1
        ttk.Label(body, text="Social links", style="CardTitle.TLabel").grid(row=row, column=0, sticky="w")
        ttk.Label(body, text="Label", style="Surface.TLabel").grid(row=row, column=1, sticky="w")
        ttk.Label(body, text="URL", style="Surface.TLabel").grid(row=row, column=2, columnspan=2, sticky="w")

        for index in range(5):
            row += 1
            ttk.Label(body, text=f"Link {index + 1}:", style="Surface.TLabel").grid(row=row, column=0, sticky="w", pady=(8, 0))
            ttk.Entry(body, textvariable=self.link_label_vars[index]).grid(row=row, column=1, sticky="ew", padx=(8, 16), pady=(8, 0))
            ttk.Entry(body, textvariable=self.link_url_vars[index]).grid(row=row, column=2, columnspan=2, sticky="ew", padx=(8, 0), pady=(8, 0))

        actions = ttk.Frame(self, style="Surface.TFrame", padding=(18, 0, 18, 18))
        actions.grid(row=2, column=0, sticky="ew")
        ttk.Button(actions, text="Reload from file", command=self.load_profile).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(actions, text="Open JSON", command=lambda: open_path(self.app.about_profile_path)).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(actions, text="Open image", command=self.app._open_about_image_file).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(actions, text="Save profile", style="Primary.TButton", command=self.save_profile).grid(row=0, column=3)

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.apply_theme(palette)
        self.load_profile()

    def _browse_image(self) -> None:
        initial = self.image_path_var.get().strip() or str(self.app.about_profile_path.parent)
        file_path = filedialog.askopenfilename(
            title="Select profile image",
            initialdir=str(Path(initial).expanduser().parent if Path(initial).expanduser().exists() else self.app.about_profile_path.parent),
            filetypes=[("Images", "*.png *.jpg *.jpeg *.webp *.gif *.bmp"), ("All files", "*.*")],
        )
        if file_path:
            try:
                relative = Path(file_path).resolve().relative_to(self.app.about_profile_path.parent.resolve())
                self.image_path_var.set(str(relative))
            except Exception:
                self.image_path_var.set(file_path)

    def load_profile(self) -> None:
        profile = load_about_profile(self.app.about_profile_path)
        self.name_var.set(str(profile.get("name", "")))
        self.title_var.set(str(profile.get("title", "")))
        self.subtitle_var.set(str(profile.get("subtitle", "")))
        self.company_var.set(str(profile.get("company", "")))
        self.project_var.set(str(profile.get("project", "")))
        self.email_var.set(str(profile.get("email", "")))
        self.handle_var.set(str(profile.get("handle", "")))
        self.image_path_var.set(str(profile.get("image_path", "")))
        self.feedback_url_var.set(str(profile.get("feedback_url", "")))
        self.contribute_url_var.set(str(profile.get("contribute_url", "")))
        self.bio_text.delete("1.0", tk.END)
        self.bio_text.insert("1.0", str(profile.get("bio", "")))
        links = profile.get("links", []) if isinstance(profile.get("links"), list) else []
        while len(links) < 5:
            links.append({"label": "", "url": ""})
        for index in range(5):
            item = links[index] if index < len(links) and isinstance(links[index], dict) else {"label": "", "url": ""}
            self.link_label_vars[index].set(str(item.get("label", "")))
            self.link_url_vars[index].set(str(item.get("url", "")))

    def save_profile(self) -> None:
        payload = {
            "name": self.name_var.get().strip(),
            "title": self.title_var.get().strip(),
            "subtitle": self.subtitle_var.get().strip(),
            "company": self.company_var.get().strip(),
            "project": self.project_var.get().strip(),
            "email": self.email_var.get().strip(),
            "handle": self.handle_var.get().strip(),
            "bio": self.bio_text.get("1.0", tk.END).strip(),
            "image_path": self.image_path_var.get().strip(),
            "feedback_url": self.feedback_url_var.get().strip(),
            "contribute_url": self.contribute_url_var.get().strip(),
            "links": [],
        }
        for label_var, url_var in zip(self.link_label_vars, self.link_url_vars):
            label = label_var.get().strip()
            url = url_var.get().strip()
            if label or url:
                payload["links"].append({"label": label or "Link", "url": url})
        self.app.about_profile_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        self.app._refresh_about_profile()
        self.app.status_var.set("About profile saved.")
        messagebox.showinfo("About profile", "Profile saved successfully.")

    def apply_theme(self, palette: ThemePalette) -> None:
        self.palette = palette
        apply_ttk_theme(self, palette)
        apply_text_widget_theme(self.bio_text, palette)




class GokulOmniConvertLiteApp(tk.Tk):
    def __init__(self, *, skip_startup_overlays: bool = False) -> None:
        super().__init__()
        self.skip_startup_overlays = skip_startup_overlays
        self.withdraw()

        self.state_store = AppStateStore()
        ensure_install_date(self.state_store.state)
        self.state_store.save()

        self.notes_path = Path(__file__).with_name("footer_notes.md")
        self.about_profile_path = ABOUT_PROFILE_PATH
        self.build_notes_path = Path(__file__).with_name("installer") / "BUILDING.md"
        self.static_about_profile_path = Path(__file__).with_name("installer") / "about_static.json"
        self.about_profile = load_about_profile(self.about_profile_path)
        self.about_photo: ImageTk.PhotoImage | None = None

        self.title(APP_NAME)
        self.geometry("1340x860")
        self.minsize(1120, 760)
        saved_geometry = str(self.state_store.get("window_geometry", "")).strip()
        if saved_geometry:
            try:
                self.geometry(saved_geometry)
            except Exception:
                pass

        self.selected_files: list[Path] = []
        self.pdf_tool_files: list[Path] = []
        self.worker_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.running = False
        self.active_run_kind = ""
        self.notes_window: MarkdownNotesWindow | None = None
        self.smtp_window: SMTPSettingsWindow | None = None
        self.build_center_window: BuildCenterWindow | None = None
        self.about_editor_window: AboutProfileEditorWindow | None = None
        self.command_palette_window: CommandPaletteWindow | None = None
        self.organizer_panel: PageOrganizerPanel | None = None
        self.splash_window: FirstLaunchSplashWindow | None = None
        self.login_popup_window: LoginReminderToast | None = None
        self.menus: list[tk.Menu] = []
        self.nav_buttons: dict[str, ttk.Button] = {}
        self.pages: dict[str, ttk.Frame] = {}
        self.current_page = "home"
        self.home_history_item_ids: dict[str, dict[str, object]] = {}
        self.history_item_ids: dict[str, dict[str, object]] = {}
        self.recent_output_item_ids: dict[str, str] = {}
        self.failed_job_item_ids: dict[str, dict[str, object]] = {}
        self.last_outputs: list[Path] = []
        self.last_output_dir: Path | None = None
        self.last_job_label = ""
        self.last_job_record: dict[str, object] | None = None
        self.last_run_origin = ""
        self.automation_after_id: str | None = None
        self.automation_active = False
        self.automation_current_files: list[Path] = []
        self.automation_current_fingerprints: list[str] = []

        self.link_status_items: dict[str, dict[str, object]] = {}
        self.link_downloaded_files: list[Path] = []
        self.link_cancel_event = threading.Event()
        self.link_pause_event = threading.Event()
        self.session_temp_root = Path(tempfile.mkdtemp(prefix="gokul_omni_convert_lite_"))
        self._tooltips: list[Tooltip] = []


        saved_output_dir = self.state_store.get("output_dir", str(Path.cwd() / "converted_output"))

        self.mode_var = tk.StringVar(value=MODE_ANY_TO_PDF)
        self.output_dir_var = tk.StringVar(value=saved_output_dir)
        self.merge_var = tk.BooleanVar(value=False)
        self.output_name_var = tk.StringVar(value=default_merged_name(MODE_ANY_TO_PDF))
        self.recursive_var = tk.BooleanVar(value=bool(self.state_store.get("recursive_scan", True)))
        self.image_format_var = tk.StringVar(value="png")
        self.image_scale_var = tk.StringVar(value="2.0")
        self.status_var = tk.StringVar(value="Ready.")
        self.mode_help_var = tk.StringVar(value=MODE_HELP[MODE_ANY_TO_PDF])
        self.supported_var = tk.StringVar(value=self._build_supported_text(MODE_ANY_TO_PDF))
        self.route_preview_var = tk.StringVar(value="Routing preview: add files to see the exact pure Python or LibreOffice plan for each item.")
        self.dependency_var = tk.StringVar(value="")
        self.home_selected_count_var = tk.StringVar(value="0 files selected")
        self.home_mode_var = tk.StringVar(value=MODE_ANY_TO_PDF)
        self.home_output_var = tk.StringVar(value=str(saved_output_dir))
        self.home_hint_var = tk.StringVar(value="Pick a mode, add files or folders, then run the batch.")
        self.theme_choice_var = tk.StringVar(value=self.state_store.get("theme", "dark"))
        self.engine_mode_var = tk.StringVar(value=str(self.state_store.get("conversion_engine", ENGINE_AUTO)))
        self.soffice_path_var = tk.StringVar(value=str(self.state_store.get("soffice_path", "")))
        self.engine_help_var = tk.StringVar(value=ENGINE_HELP.get(str(self.state_store.get("conversion_engine", ENGINE_AUTO)), ""))
        self.active_engine_var = tk.StringVar(value="")
        self.history_detail_var = tk.StringVar(value="Select a recent job to inspect details or re-use settings.")
        self.pdf_tool_var = tk.StringVar(value=PDF_TOOL_MERGE)
        self.pdf_tool_help_var = tk.StringVar(value=PDF_TOOL_HELP[PDF_TOOL_MERGE])
        self.pdf_tool_output_name_var = tk.StringVar(value="merged_pdfs")
        self.pdf_tool_page_spec_var = tk.StringVar(value="")
        self.pdf_tool_every_n_var = tk.StringVar(value="2")
        self.pdf_tool_watermark_text_var = tk.StringVar(value="CONFIDENTIAL")
        self.pdf_tool_watermark_image_var = tk.StringVar(value="")
        self.pdf_tool_font_size_var = tk.StringVar(value="42")
        self.pdf_tool_rotation_var = tk.StringVar(value="45")
        self.pdf_tool_opacity_var = tk.StringVar(value="0.18")
        self.pdf_tool_position_var = tk.StringVar(value="center")
        self.pdf_tool_image_scale_var = tk.StringVar(value="40")
        self.pdf_tool_metadata_title_var = tk.StringVar(value="")
        self.pdf_tool_metadata_author_var = tk.StringVar(value="")
        self.pdf_tool_metadata_subject_var = tk.StringVar(value="")
        self.pdf_tool_metadata_keywords_var = tk.StringVar(value="")
        self.pdf_tool_metadata_clear_var = tk.BooleanVar(value=False)
        self.pdf_tool_password_var = tk.StringVar(value="")
        self.pdf_tool_owner_password_var = tk.StringVar(value="")
        self.pdf_tool_compression_profile_var = tk.StringVar(value="balanced")
        self.pdf_tool_redact_rect_var = tk.StringVar(value="10%,10%,90%,25%")
        self.pdf_tool_replacement_text_var = tk.StringVar(value="")
        self.pdf_tool_hint_var = tk.StringVar(value=self._build_pdf_tool_hint(PDF_TOOL_MERGE))
        self.ocr_input_var = tk.StringVar(value="")
        default_ocr_output = str(self.state_store.get("ocr_output_dir", "") or (Path(saved_output_dir) / "ocr"))
        self.ocr_output_var = tk.StringVar(value=default_ocr_output)
        self.ocr_language_var = tk.StringVar(value=str(self.state_store.get("ocr_language", "eng")))
        self.ocr_dpi_var = tk.StringVar(value=str(self.state_store.get("ocr_dpi", 220)))
        self.ocr_psm_var = tk.StringVar(value=str(self.state_store.get("ocr_psm", 6)))
        self.tesseract_path_var = tk.StringVar(value=str(self.state_store.get("tesseract_path", "")))
        self.ocr_status_var = tk.StringVar(value="OCR tools are ready. Add an image or PDF to begin.")
        self.ocr_dependency_var = tk.StringVar(value="")
        self.preset_name_var = tk.StringVar(value="My preset")
        self.watch_source_dir_var = tk.StringVar(value="")
        self.watch_output_dir_var = tk.StringVar(value=saved_output_dir)
        self.watch_mode_var = tk.StringVar(value=MODE_ANY_TO_PDF)
        self.watch_merge_var = tk.BooleanVar(value=False)
        self.watch_recursive_var = tk.BooleanVar(value=bool(self.state_store.get("recursive_scan", True)))
        self.watch_interval_var = tk.StringVar(value="15")
        self.watch_output_name_var = tk.StringVar(value="watch_output")
        self.watch_engine_var = tk.StringVar(value=str(self.state_store.get("conversion_engine", ENGINE_AUTO)))
        self.watch_archive_var = tk.BooleanVar(value=False)
        self.watch_archive_dir_var = tk.StringVar(value="")
        self.watch_zip_var = tk.BooleanVar(value=False)
        self.watch_report_var = tk.BooleanVar(value=True)
        self.watch_mail_var = tk.BooleanVar(value=False)
        self.watch_skip_existing_var = tk.BooleanVar(value=True)
        self.watch_status_var = tk.StringVar(value="Automation is idle.")
        self.watch_summary_var = tk.StringVar(value="Configure a watch folder, save presets, and bundle outputs from this page.")
        self.share_bundle_name_var = tk.StringVar(value="gokul_outputs_bundle")
        self.share_report_name_var = tk.StringVar(value="gokul_last_run_report")

        self.link_cache_dir_var = tk.StringVar(value=str(self.state_store.get("link_cache_dir", "")))
        self.link_timeout_var = tk.StringVar(value=str(self.state_store.get("link_timeout", 25)))
        self.link_keep_downloads_var = tk.BooleanVar(value=bool(self.state_store.get("link_keep_downloads", True)))
        self.link_cache_max_age_var = tk.StringVar(value=str(self.state_store.get("link_cache_max_age_days", 30)))
        self.link_cache_max_size_var = tk.StringVar(value=str(self.state_store.get("link_cache_max_size_mb", 512)))
        self.link_cache_summary_var = tk.StringVar(value="Cache stats are not loaded yet.")
        self.link_status_summary_var = tk.StringVar(value="Paste one or more direct file links or page URLs, then fetch them into the same conversion queue.")
        self.link_recent_summary_var = tk.StringVar(value="")
        self.link_fetch_count_var = tk.StringVar(value="0 downloaded")
        self.link_auto_start_pending = False

        self.performance_mode_var = tk.StringVar(value=str(self.state_store.get("performance_mode", "balanced")))
        self.favorite_preset_summary_var = tk.StringVar(value="Favorite presets will appear here after you star them.")

        self.splash_enabled_var = tk.BooleanVar(value=bool(self.state_store.get("splash_enabled", True)))
        self.splash_gif_path_var = tk.StringVar(value=str(self.state_store.get("splash_gif_path", "assets/gokul_splash.gif")))
        self.login_popup_enabled_var = tk.BooleanVar(value=bool(self.state_store.get("login_popup_enabled", True)))
        self.install_date_var = tk.StringVar(value=str(self.state_store.get("install_date", "")))
        self.login_popup_state_var = tk.StringVar(value="")
        self.auto_open_output_var = tk.BooleanVar(value=bool(self.state_store.get("auto_open_output_folder", False)))
        self.restore_session_var = tk.BooleanVar(value=bool(self.state_store.get("restore_last_session", True)))
        self.cleanup_temp_var = tk.BooleanVar(value=bool(self.state_store.get("cleanup_temp_on_exit", True)))
        self.update_checker_enabled_var = tk.BooleanVar(value=bool(self.state_store.get("update_checker_enabled", False)))
        self.last_update_check_var = tk.StringVar(value=str(self.state_store.get("last_update_check", "")))

        smtp_config = SMTPSettings.from_dict(self.state_store.get("smtp_settings", {}))
        self.smtp_host_var = tk.StringVar(value=smtp_config.host)
        self.smtp_port_var = tk.StringVar(value=str(smtp_config.port))
        self.smtp_username_var = tk.StringVar(value=smtp_config.username)
        self.smtp_password_var = tk.StringVar(value=smtp_config.password)
        self.smtp_sender_var = tk.StringVar(value=smtp_config.sender or str(self.about_profile.get("email", "")).strip())
        self.smtp_default_to_var = tk.StringVar(value=smtp_config.default_to)
        self.smtp_use_ssl_var = tk.BooleanVar(value=smtp_config.use_ssl)
        self.smtp_use_starttls_var = tk.BooleanVar(value=smtp_config.use_starttls)
        self.smtp_save_password_var = tk.BooleanVar(value=smtp_config.save_password)
        self.smtp_status_var = tk.StringVar(value="Configure SMTP here when you want direct send instead of a draft.")

        watch_config = normalize_watch_config(self.state_store.watch_config())
        self.watch_source_dir_var.set(watch_config.source_dir)
        self.watch_output_dir_var.set(watch_config.output_dir or saved_output_dir)
        self.watch_mode_var.set(watch_config.mode)
        self.watch_merge_var.set(watch_config.merge_to_one_pdf)
        self.watch_recursive_var.set(watch_config.recursive)
        self.watch_interval_var.set(str(watch_config.interval_seconds))
        self.watch_output_name_var.set(watch_config.merged_output_name)
        self.watch_engine_var.set(watch_config.engine_mode)
        self.watch_archive_var.set(watch_config.archive_processed)
        self.watch_archive_dir_var.set(watch_config.archive_dir)
        self.watch_zip_var.set(watch_config.create_zip_bundle)
        self.watch_report_var.set(watch_config.create_report)
        self.watch_mail_var.set(watch_config.open_mail_draft)
        self.watch_skip_existing_var.set(watch_config.skip_existing_on_start)

        self._build_shell()
        self._build_menu()
        self._refresh_dependency_status()
        self._refresh_about_profile()
        self._update_mode_controls()
        self._update_pdf_tool_controls()
        self._refresh_history_views()
        self._apply_theme(initial=True)
        startup_page = str(self.state_store.get("last_page", "home")).strip() or "home"
        self._show_page(startup_page if startup_page in self.pages else "home")
        self._update_login_popup_state()
        self._refresh_recent_links_summary()
        self._refresh_link_cache_summary()
        self._sync_static_about_profile()
        self._restore_last_session_snapshot()
        self._refresh_favorite_preset_widgets()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(120, self._finish_startup)


    def _finish_startup(self) -> None:
        self._update_login_popup_state()
        if self.skip_startup_overlays:
            self.deiconify()
            self.lift()
            return

        if should_show_first_launch_splash(dict(self.state_store.state)) and bool(self.splash_enabled_var.get()):
            self.splash_window = FirstLaunchSplashWindow(
                self,
                gif_path=self._resolve_splash_gif_path(),
                palette=self.palette,
                on_close=self._on_splash_closed,
            )
            return

        self.deiconify()
        self.lift()
        self.after(700, self._maybe_show_login_popup)

    def _on_splash_closed(self) -> None:
        self.splash_window = None
        self.state_store.update(
            splash_seen=True,
            splash_enabled=bool(self.splash_enabled_var.get()),
            splash_gif_path=self.splash_gif_path_var.get().strip() or "assets/gokul_splash.gif",
        )
        self.install_date_var.set(str(self.state_store.get("install_date", "")))
        self.deiconify()
        self.lift()
        self.after(700, self._maybe_show_login_popup)

    def _maybe_show_login_popup(self) -> None:
        self._update_login_popup_state()
        if self.skip_startup_overlays:
            return
        if self.login_popup_window and self.login_popup_window.winfo_exists():
            return
        snapshot = dict(self.state_store.state)
        snapshot["login_popup_enabled"] = bool(self.login_popup_enabled_var.get())
        if should_show_login_popup(snapshot):
            self.state_store.update(
                login_popup_enabled=bool(self.login_popup_enabled_var.get()),
                login_popup_last_shown=iso_now(),
            )
            self.login_popup_window = LoginReminderToast(
                self,
                palette=self.palette,
                on_dismiss=self._dismiss_login_popup,
                on_complete=self._complete_login_popup,
            )
            self._update_login_popup_state()

    def _dismiss_login_popup(self) -> None:
        self.login_popup_window = None
        self.state_store.update(
            login_popup_dismissed=True,
            login_popup_completed=False,
            login_popup_enabled=bool(self.login_popup_enabled_var.get()),
            login_popup_last_shown=iso_now(),
        )
        self._update_login_popup_state()
        self.status_var.set("Login reminder dismissed permanently.")

    def _complete_login_popup(self) -> None:
        self.login_popup_window = None
        self.state_store.update(
            login_popup_completed=True,
            login_popup_dismissed=False,
            login_popup_enabled=bool(self.login_popup_enabled_var.get()),
            login_popup_last_shown=iso_now(),
        )
        self._update_login_popup_state()
        self.status_var.set("Login reminder marked as completed.")

    def _reset_login_popup_state(self) -> None:
        self.state_store.update(
            login_popup_dismissed=False,
            login_popup_completed=False,
            login_popup_last_shown="",
            login_popup_enabled=bool(self.login_popup_enabled_var.get()),
        )
        self._update_login_popup_state()
        self.status_var.set("Login reminder state reset.")

    def _update_login_popup_state(self) -> None:
        self.install_date_var.set(str(self.state_store.get("install_date", "")))
        snapshot = dict(self.state_store.state)
        snapshot["login_popup_enabled"] = bool(self.login_popup_enabled_var.get())
        self.login_popup_state_var.set(summarize_login_popup_state(snapshot))

    def _resolve_splash_gif_path(self) -> Path:
        value = self.splash_gif_path_var.get().strip() or "assets/gokul_splash.gif"
        path = Path(value).expanduser()
        if not path.is_absolute():
            path = Path(__file__).resolve().parent / path
        return path

    def _preview_splash(self) -> None:
        if self.splash_window and self.splash_window.winfo_exists():
            self.splash_window.lift()
            return
        self.splash_window = FirstLaunchSplashWindow(
            self,
            gif_path=self._resolve_splash_gif_path(),
            palette=self.palette,
            on_close=self._on_preview_splash_closed,
        )

    def _on_preview_splash_closed(self) -> None:
        self.splash_window = None
        self.lift()

    def _browse_splash_gif_path(self) -> None:
        initial = self._resolve_splash_gif_path()
        file_path = filedialog.askopenfilename(
            title="Select splash GIF",
            initialdir=str(initial.parent),
            filetypes=[("GIF files", "*.gif"), ("All files", "*.*")],
        )
        if file_path:
            try:
                relative = Path(file_path).resolve().relative_to(Path(__file__).resolve().parent)
                self.splash_gif_path_var.set(str(relative))
            except Exception:
                self.splash_gif_path_var.set(file_path)
            self.status_var.set("Splash asset path updated.")
            self._persist_state()

    def _clear_splash_gif_path(self) -> None:
        self.splash_gif_path_var.set("assets/gokul_splash.gif")
        self._persist_state()
        self.status_var.set("Splash asset path reset to the bundled placeholder GIF.")

    def _clear_soffice_path(self) -> None:
        self.soffice_path_var.set("")
        self._persist_state()
        self.status_var.set("LibreOffice path cleared.")

    def _sync_static_about_profile(self) -> None:
        try:
            self.static_about_profile_path.parent.mkdir(parents=True, exist_ok=True)
            snapshot = {
                "name": str(self.about_profile.get("name", APP_NAME)),
                "title": str(self.about_profile.get("title", "")),
                "subtitle": str(self.about_profile.get("subtitle", "")),
                "company": str(self.about_profile.get("company", "")),
                "project": str(self.about_profile.get("project", "")),
                "email": str(self.about_profile.get("email", "")),
                "handle": str(self.about_profile.get("handle", "")),
                "feedback_url": str(self.about_profile.get("feedback_url", "")),
                "contribute_url": str(self.about_profile.get("contribute_url", "")),
                "image_path": str(self.about_profile.get("image_path", "")),
            }
            self.static_about_profile_path.write_text(json.dumps(snapshot, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _build_shell(self) -> None:
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.header = ttk.Frame(self, padding=(22, 18, 22, 14))
        self.header.grid(row=0, column=0, sticky="ew")
        self.header.grid_columnconfigure(0, weight=1)

        ttk.Label(self.header, text=APP_NAME, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            self.header,
            text="A cleaner desktop workspace for batch conversion, PDF tools, OCR, a visual page organizer, automation presets, watch folders, share bundles, password protection, compression, and pure Python conversion with optional LibreOffice when you choose it.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        header_actions = ttk.Frame(self.header)
        header_actions.grid(row=0, column=1, rowspan=2, sticky="e")
        ttk.Label(header_actions, text="Theme:").grid(row=0, column=0, sticky="e", padx=(0, 8))
        self.theme_combo = ttk.Combobox(
            header_actions,
            textvariable=self.theme_choice_var,
            values=["dark", "light", "system"],
            width=10,
            state="readonly",
        )
        self.theme_combo.grid(row=0, column=1, sticky="e")
        self.theme_combo.bind("<<ComboboxSelected>>", lambda _event: self._apply_theme())
        ttk.Button(header_actions, text="Open Notes", command=self._open_notes_window).grid(row=0, column=2, padx=(10, 0))
        ttk.Button(header_actions, text="Mail", command=self._open_smtp_window).grid(row=0, column=3, padx=(10, 0))
        ttk.Button(header_actions, text="Build", command=self._open_build_center_window).grid(row=0, column=4, padx=(10, 0))
        ttk.Button(header_actions, text="About", command=self._show_about).grid(row=0, column=5, padx=(10, 0))
        ttk.Button(header_actions, text="OCR", command=lambda: self._show_page("ocr")).grid(row=0, column=6, padx=(10, 0))
        ttk.Button(header_actions, text="Organizer", command=lambda: self._show_page("organizer")).grid(row=0, column=7, padx=(10, 0))
        ttk.Button(header_actions, text="Automation", command=lambda: self._show_page("automation")).grid(row=0, column=8, padx=(10, 0))
        ttk.Button(header_actions, text="Go to Convert", style="Primary.TButton", command=lambda: self._show_page("convert")).grid(
            row=0, column=9, padx=(10, 0)
        )

        self.body = ttk.Frame(self)
        self.body.grid(row=1, column=0, sticky="nsew", padx=22, pady=(0, 10))
        self.body.grid_columnconfigure(1, weight=1)
        self.body.grid_rowconfigure(0, weight=1)

        self.sidebar = ttk.Frame(self.body, style="Sidebar.TFrame", padding=16)
        self.sidebar.grid(row=0, column=0, sticky="nsw", padx=(0, 14))
        self.sidebar.grid_columnconfigure(0, weight=1)
        ttk.Label(self.sidebar, text="Workspace", style="Sidebar.TLabel", font=("Segoe UI", 11, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 10)
        )

        for index, (page_name, label) in enumerate(NAV_PAGES, start=1):
            button = ttk.Button(
                self.sidebar,
                text=label,
                style="Nav.TButton",
                command=lambda name=page_name: self._show_page(name),
            )
            button.grid(row=index, column=0, sticky="ew", pady=4)
            self.nav_buttons[page_name] = button

        ttk.Separator(self.sidebar, orient="horizontal").grid(row=len(NAV_PAGES) + 1, column=0, sticky="ew", pady=14)
        ttk.Label(
            self.sidebar,
            text="Current status",
            style="SidebarMuted.TLabel",
            font=("Segoe UI", 10, "bold"),
        ).grid(row=len(NAV_PAGES) + 2, column=0, sticky="w")
        self.sidebar_status = ttk.Label(self.sidebar, textvariable=self.status_var, style="SidebarMuted.TLabel", wraplength=180, justify="left")
        self.sidebar_status.grid(row=len(NAV_PAGES) + 3, column=0, sticky="ew", pady=(6, 0))

        self.content = ttk.Frame(self.body)
        self.content.grid(row=0, column=1, sticky="nsew")
        self.content.grid_rowconfigure(0, weight=1)
        self.content.grid_columnconfigure(0, weight=1)

        self._build_home_page()
        self._build_convert_page()
        self._build_pdf_tools_page()
        self._build_ocr_page()
        self._build_organizer_page()
        self._build_automation_page()
        self._build_history_page()
        self._build_settings_page()
        self._build_about_page()

        self.footer = ttk.Frame(self, padding=(22, 12, 22, 18))
        self.footer.grid(row=2, column=0, sticky="ew")
        self.footer.grid_columnconfigure(0, weight=1)
        ttk.Label(self.footer, text=f"{APP_NAME}  |  {APP_VERSION}").grid(row=0, column=0, sticky="w")
        footer_actions = ttk.Frame(self.footer)
        footer_actions.grid(row=0, column=1, sticky="e")
        ttk.Button(footer_actions, text="About", command=self._show_about).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(footer_actions, text="Mail", command=self._open_smtp_window).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(footer_actions, text="Build Center", command=self._open_build_center_window).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(footer_actions, text="Open footer notes", command=self._open_notes_window).grid(row=0, column=3)
        ttk.Label(self.footer, textvariable=self.status_var).grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))

    def _build_menu(self) -> None:
        menu_bar = tk.Menu(self)

        file_menu = tk.Menu(menu_bar)
        file_menu.add_command(label="Add Files...", command=self._add_files, accelerator="Ctrl+O")
        file_menu.add_command(label="Add Folder...", command=self._add_folder, accelerator="Ctrl+Shift+O")
        file_menu.add_separator()
        file_menu.add_command(label="Choose Output Folder...", command=self._browse_output_dir)
        file_menu.add_command(label="Open Output Folder", command=self._open_output_folder)
        file_menu.add_command(label="Create ZIP Bundle for Last Outputs", command=self._create_zip_bundle_for_last_outputs)
        file_menu.add_command(label="Export Last Run Report", command=self._export_last_run_report)
        file_menu.add_command(label="Create Mail Draft for Last Outputs", command=self._open_mail_draft_for_last_outputs)
        file_menu.add_command(label="Save EML Draft for Last Outputs", command=self._save_eml_draft_for_last_outputs)
        file_menu.add_command(label="Send Last Outputs via SMTP", command=self._send_last_outputs_via_smtp)
        file_menu.add_separator()
        file_menu.add_command(label="Export Settings Snapshot", command=self._export_state_snapshot_action)
        file_menu.add_command(label="Import Settings Snapshot", command=self._import_state_snapshot_action)
        file_menu.add_command(label="Export Diagnostics JSON", command=self._export_diagnostics_report_action)
        file_menu.add_command(label="Export App Logs", command=self._export_current_logs_action)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._on_close)
        menu_bar.add_cascade(label="File", menu=file_menu)

        convert_menu = tk.Menu(menu_bar)
        convert_menu.add_command(label="Start Conversion", command=self._start_conversion, accelerator="Ctrl+R")
        convert_menu.add_command(label="Refresh Dependency Status", command=self._refresh_dependency_status)
        convert_menu.add_separator()
        convert_menu.add_command(label="Clear Inputs", command=self._clear_inputs)
        convert_menu.add_command(label="Clear Activity Log", command=self._clear_log, accelerator="Ctrl+L")
        menu_bar.add_cascade(label="Convert", menu=convert_menu)

        pdf_menu = tk.Menu(menu_bar)
        pdf_menu.add_command(label="Open PDF Tools", command=lambda: self._show_page("pdf_tools"))
        pdf_menu.add_command(label="Open Organizer", command=lambda: self._show_page("organizer"))
        pdf_menu.add_command(label="Load PDF in Organizer...", command=self._open_pdf_in_organizer)
        pdf_menu.add_command(label="Add PDF Files...", command=self._add_pdf_tool_files)
        pdf_menu.add_command(label="Add PDF Folder...", command=self._add_pdf_tool_folder)
        pdf_menu.add_separator()
        pdf_menu.add_command(label="Run Current PDF Tool", command=self._start_pdf_tool)
        pdf_menu.add_command(label="Clear PDF Tool Inputs", command=self._clear_pdf_tool_inputs)
        menu_bar.add_cascade(label="PDF Tools", menu=pdf_menu)

        ocr_menu = tk.Menu(menu_bar)
        ocr_menu.add_command(label="Open OCR", command=lambda: self._show_page("ocr"))
        ocr_menu.add_command(label="Choose OCR Input...", command=self._browse_ocr_input)
        ocr_menu.add_command(label="Choose OCR Output Folder...", command=self._browse_ocr_output)
        ocr_menu.add_separator()
        ocr_menu.add_command(label="Image to Searchable PDF", command=self._start_ocr_image_pdf)
        ocr_menu.add_command(label="PDF to Searchable PDF", command=self._start_ocr_pdf_pdf)
        ocr_menu.add_command(label="Extract OCR Text", command=self._start_ocr_text)
        ocr_menu.add_separator()
        ocr_menu.add_command(label="Test Tesseract Path", command=self._test_tesseract_path)
        menu_bar.add_cascade(label="OCR", menu=ocr_menu)

        automation_menu = tk.Menu(menu_bar)
        automation_menu.add_command(label="Open Automation", command=lambda: self._show_page("automation"))
        automation_menu.add_command(label="Save Current Preset", command=self._save_current_preset)
        automation_menu.add_command(label="Scan Watch Folder Now", command=self._scan_watch_folder_now)
        automation_menu.add_command(label="Start Watcher", command=self._start_watch_automation)
        automation_menu.add_command(label="Stop Watcher", command=self._stop_watch_automation)
        automation_menu.add_separator()
        automation_menu.add_command(label="Create ZIP Bundle for Last Outputs", command=self._create_zip_bundle_for_last_outputs)
        automation_menu.add_command(label="Export Last Run Report", command=self._export_last_run_report)
        automation_menu.add_command(label="Send Last Outputs via SMTP", command=self._send_last_outputs_via_smtp)
        menu_bar.add_cascade(label="Automation", menu=automation_menu)

        view_menu = tk.Menu(menu_bar)
        for page_name, label in NAV_PAGES:
            view_menu.add_command(label=label, command=lambda name=page_name: self._show_page(name))
        view_menu.add_separator()
        theme_menu = tk.Menu(view_menu)
        for choice in ("dark", "light", "system"):
            theme_menu.add_command(label=choice.title(), command=lambda value=choice: self._set_theme(value))
        view_menu.add_cascade(label="Theme", menu=theme_menu)
        menu_bar.add_cascade(label="View", menu=view_menu)

        help_menu = tk.Menu(menu_bar)
        help_menu.add_command(label="Open Footer Notes", command=self._open_notes_window)
        help_menu.add_command(label="Open Organizer", command=lambda: self._show_page("organizer"))
        help_menu.add_command(label="Open OCR", command=lambda: self._show_page("ocr"))
        help_menu.add_command(label="Open Markdown File", command=lambda: open_path(self.notes_path))
        help_menu.add_command(label="Open SMTP Delivery", command=self._open_smtp_window)
        help_menu.add_command(label="Open Build Center", command=self._open_build_center_window)
        help_menu.add_command(label="Quick actions", command=self._open_command_palette)
        help_menu.add_command(label="Check for updates (placeholder)", command=self._check_for_updates_placeholder)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self._show_about)
        menu_bar.add_cascade(label="Help", menu=help_menu)

        self.config(menu=menu_bar)
        self.menus = [menu_bar, file_menu, convert_menu, pdf_menu, ocr_menu, automation_menu, view_menu, theme_menu, help_menu]

        self.bind_all("<Control-o>", lambda _event: self._add_files())
        self.bind_all("<Control-O>", lambda _event: self._add_folder())
        self.bind_all("<Control-r>", lambda _event: self._start_conversion())
        self.bind_all("<Control-Return>", lambda _event: self._start_conversion())
        self.bind_all("<Control-Shift-Return>", lambda _event: self._start_pdf_tool())
        self.bind_all("<Control-l>", lambda _event: self._clear_log())
        self.bind_all("<Control-Shift-L>", lambda _event: self._focus_link_input())
        self.bind_all("<Control-k>", lambda _event: self._open_command_palette())
        self.bind_all("<Control-comma>", lambda _event: self._show_page("settings"))
        self.bind_all("<F5>", lambda _event: self._refresh_dependency_status())

    def _build_home_page(self) -> None:
        page = ttk.Frame(self.content)
        page.grid(row=0, column=0, sticky="nsew")
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(2, weight=1)
        self.pages["home"] = page

        hero = ttk.Frame(page, style="Card.TFrame", padding=22)
        hero.grid(row=0, column=0, sticky="ew")
        hero.grid_columnconfigure(0, weight=1)

        ttk.Label(hero, text="Start from one clean workspace", style="HeroTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            hero,
            text="Use the convert workspace for pure-Python-first conversions, batch output control, and quick access to related tools such as organizer, automation, OCR, and output sharing.",
            style="HeroBody.TLabel",
            wraplength=880,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(8, 12))

        actions = ttk.Frame(hero, style="Surface.TFrame")
        actions.grid(row=0, column=1, rowspan=2, sticky="e")
        ttk.Button(actions, text="Add Files", command=self._add_files).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(actions, text="Add Folder", command=self._add_folder).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(actions, text="Organizer", command=lambda: self._show_page("organizer")).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(actions, text="Start Batch", style="Primary.TButton", command=lambda: (self._show_page("convert"), self._start_conversion())).grid(
            row=0, column=3
        )

        metrics = ttk.Frame(page)
        metrics.grid(row=1, column=0, sticky="ew", pady=(14, 14))
        for column in range(3):
            metrics.grid_columnconfigure(column, weight=1)

        self._create_metric_card(metrics, 0, "Selected inputs", self.home_selected_count_var)
        self._create_metric_card(metrics, 1, "Current mode", self.home_mode_var)
        self._create_metric_card(metrics, 2, "Output folder", self.home_output_var)

        lower = ttk.Frame(page)
        lower.grid(row=2, column=0, sticky="nsew")
        lower.grid_columnconfigure(0, weight=2)
        lower.grid_columnconfigure(1, weight=1)
        lower.grid_rowconfigure(0, weight=1)

        history_card = ttk.LabelFrame(lower, text="Recent jobs")
        history_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        history_card.grid_columnconfigure(0, weight=1)
        history_card.grid_rowconfigure(1, weight=1)

        ttk.Label(
            history_card,
            text="The most recent jobs are stored locally so you can inspect what ran last and reuse the same settings.",
            style="CardBody.TLabel",
            wraplength=660,
            justify="left",
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        self.home_history_tree = ttk.Treeview(
            history_card,
            columns=("time", "status", "mode", "files"),
            show="headings",
            height=9,
        )
        for col, title, width in (
            ("time", "Time", 170),
            ("status", "Status", 110),
            ("mode", "Mode", 360),
            ("files", "Files", 80),
        ):
            self.home_history_tree.heading(col, text=title)
            self.home_history_tree.column(col, width=width, anchor="w")
        self.home_history_tree.grid(row=1, column=0, sticky="nsew")
        self.home_history_tree.bind("<<TreeviewSelect>>", self._on_home_history_selected)

        dep_card = ttk.LabelFrame(lower, text="Dependency status and guidance")
        dep_card.grid(row=0, column=1, sticky="nsew")
        dep_card.grid_columnconfigure(0, weight=1)

        self.home_dependency_label = ttk.Label(dep_card, textvariable=self.dependency_var, style="CardBody.TLabel", wraplength=320, justify="left")
        self.home_dependency_label.grid(row=0, column=0, sticky="nw")
        ttk.Label(dep_card, textvariable=self.home_hint_var, style="CardBody.TLabel", wraplength=320, justify="left").grid(
            row=1, column=0, sticky="nw", pady=(12, 0)
        )
        button_row = ttk.Frame(dep_card)
        button_row.grid(row=2, column=0, sticky="ew", pady=(16, 0))
        ttk.Button(button_row, text="Refresh dependencies", command=self._refresh_dependency_status).grid(row=0, column=0, sticky="w")
        ttk.Button(button_row, text="Quick actions", command=self._open_command_palette).grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Button(dep_card, text="Open organizer", command=lambda: self._show_page("organizer")).grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Button(dep_card, text="Open footer notes", command=self._open_notes_window).grid(row=4, column=0, sticky="w", pady=(10, 0))
        ttk.Label(dep_card, text="Favorite presets", style="CardTitle.TLabel").grid(row=5, column=0, sticky="w", pady=(16, 0))
        ttk.Label(dep_card, textvariable=self.favorite_preset_summary_var, style="CardBody.TLabel", wraplength=320, justify="left").grid(
            row=6, column=0, sticky="w", pady=(6, 0)
        )
        self.home_favorite_presets_frame = ttk.Frame(dep_card, style="Surface.TFrame")
        self.home_favorite_presets_frame.grid(row=7, column=0, sticky="ew", pady=(8, 0))


def _refresh_favorite_preset_widgets(self) -> None:
    favorites = self.state_store.favorite_presets()
    if favorites:
        names = ", ".join(str(item.get("name", "")).strip() for item in favorites[:3])
        if len(favorites) > 3:
            names += f", +{len(favorites) - 3} more"
        self.favorite_preset_summary_var.set(f"Star presets in Automation, then launch them from here. Current favorites: {names}.")
    else:
        self.favorite_preset_summary_var.set("Star presets in Automation to pin your most-used workflows here.")
    frame = getattr(self, "home_favorite_presets_frame", None)
    if frame is None:
        return
    for child in frame.winfo_children():
        child.destroy()
    if not favorites:
        ttk.Label(frame, text="No favorite presets yet.", style="CardBody.TLabel").grid(row=0, column=0, sticky="w")
        return
    for index, preset in enumerate(favorites[:4]):
        button = ttk.Button(
            frame,
            text=f"★ {preset.get('name', '')}",
            command=lambda item=preset: self._apply_preset_record(item, start=False),
        )
        button.grid(row=index, column=0, sticky="ew", pady=(0 if index == 0 else 6, 0))
        self._attach_tooltip(button, f"{preset.get('mode', '')} • {preset.get('engine_mode', '')}")
    run_row = ttk.Frame(frame, style="Surface.TFrame")
    run_row.grid(row=min(len(favorites), 4), column=0, sticky="ew", pady=(10, 0))
    ttk.Button(run_row, text="Open Automation", command=lambda: self._show_page("automation")).grid(row=0, column=0, sticky="w")
    ttk.Button(run_row, text="Run first favorite", command=self._run_first_favorite_preset).grid(row=0, column=1, sticky="w", padx=(8, 0))

def _run_first_favorite_preset(self) -> None:
    favorites = self.state_store.favorite_presets()
    if not favorites:
        self.status_var.set("No favorite preset is available yet.")
        self._show_page("automation")
        return
    self._apply_preset_record(favorites[0], start=True)

def _apply_preset_record(self, preset: dict[str, object], *, start: bool = False) -> None:
    self.mode_var.set(str(preset.get("mode", MODE_ANY_TO_PDF)))
    self.output_dir_var.set(str(preset.get("output_dir", self.output_dir_var.get())))
    self.merge_var.set(bool(preset.get("merge_to_one_pdf", False)))
    self.output_name_var.set(str(preset.get("merged_output_name", default_merged_name(self.mode_var.get()))))
    self.image_format_var.set(str(preset.get("image_format", "png")))
    self.image_scale_var.set(str(preset.get("image_scale", "2.0")))
    self.engine_mode_var.set(str(preset.get("engine_mode", ENGINE_AUTO)))
    self.recursive_var.set(bool(preset.get("recursive", True)))
    self._refresh_dependency_status()
    self._update_mode_controls()
    self._show_page("convert")
    name = str(preset.get("name", "")).strip()
    self.status_var.set(f"Applied preset '{name}'.")
    if start:
        self.after(80, self._start_conversion)

def _toggle_selected_preset_favorite(self) -> None:
    preset = self._selected_preset()
    if not preset:
        messagebox.showinfo("Presets", "Select a preset first.")
        return
    name = str(preset.get("name", "")).strip()
    favorite = not bool(preset.get("favorite", False))
    self.state_store.set_preset_favorite(name, favorite)
    self._refresh_presets_view()
    self._refresh_favorite_preset_widgets()
    self._log_automation(f"{'Starred' if favorite else 'Unstarred'} preset '{name}'.")

def _focus_link_input(self) -> None:
    self._show_page("convert")
    try:
        self.link_input_text.focus_set()
        self.link_input_text.mark_set("insert", "1.0")
    except Exception:
        pass

def _build_quick_actions(self) -> list[QuickAction]:
    return [
        QuickAction("Add files", self._add_files, hint="Browse and add input files", keywords="open import files"),
        QuickAction("Add folder", self._add_folder, hint="Scan a folder for supported files", keywords="directory batch"),
        QuickAction("Start conversion", self._start_conversion, hint="Run the current Convert queue", keywords="run batch ctrl+enter"),
        QuickAction("Run PDF tool", self._start_pdf_tool, hint="Start the active PDF tool", keywords="pdf tool"),
        QuickAction("Open settings", lambda: self._show_page("settings"), hint="Jump to app settings", keywords="preferences"),
        QuickAction("Focus links box", self._focus_link_input, hint="Jump to the online URL input area", keywords="url links"),
        QuickAction("Open organizer", lambda: self._show_page("organizer"), hint="Visual page organization", keywords="pages reorder"),
        QuickAction("Open OCR", lambda: self._show_page("ocr"), hint="Searchable PDF and OCR text tools", keywords="scan text"),
        QuickAction("Open Automation", lambda: self._show_page("automation"), hint="Presets and watch folder tools", keywords="presets watch"),
        QuickAction("Open Build Center", self._open_build_center_window, hint="Diagnostics and release prep", keywords="installer diagnostics"),
        QuickAction("Open SMTP Delivery", self._open_smtp_window, hint="Draft or send outputs by email", keywords="mail email"),
        QuickAction("Check for updates", self._check_for_updates_placeholder, hint="Record a local update-check stamp", keywords="release"),
    ]

def _open_command_palette(self) -> None:
    if self.command_palette_window and self.command_palette_window.winfo_exists():
        self.command_palette_window.lift()
        return
    self.command_palette_window = CommandPaletteWindow(self, actions=self._build_quick_actions(), palette=self.palette)
    self.command_palette_window.bind("<Destroy>", lambda _event: setattr(self, "command_palette_window", None), add="+")

def _attach_tooltip(self, widget: tk.Widget, text: str) -> None:
    if not text:
        return
    self._tooltips.append(Tooltip(widget, text))


    def _build_convert_page(self) -> None:
        page = ttk.Frame(self.content)
        page.grid(row=0, column=0, sticky="nsew")
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(2, weight=1)
        self.pages["convert"] = page

        mode_card = ttk.LabelFrame(page, text="Mode and supported inputs")
        mode_card.grid(row=0, column=0, sticky="ew")
        mode_card.grid_columnconfigure(1, weight=1)

        ttk.Label(mode_card, text="Conversion mode:", style="Surface.TLabel").grid(row=0, column=0, sticky="w")
        self.mode_combo = ttk.Combobox(mode_card, textvariable=self.mode_var, values=MODE_ORDER, state="readonly", width=38)
        self.mode_combo.grid(row=0, column=1, sticky="ew", padx=(10, 10))
        self.mode_combo.bind("<<ComboboxSelected>>", self._on_mode_changed)
        ttk.Button(mode_card, text="Refresh dependencies", command=self._refresh_dependency_status).grid(row=0, column=2, sticky="e")

        ttk.Label(mode_card, textvariable=self.mode_help_var, style="CardBody.TLabel", wraplength=980, justify="left").grid(
            row=1, column=0, columnspan=3, sticky="ew", pady=(10, 6)
        )
        ttk.Label(mode_card, textvariable=self.supported_var, style="CardBody.TLabel", wraplength=980, justify="left").grid(
            row=2, column=0, columnspan=3, sticky="ew"
        )
        ttk.Label(mode_card, textvariable=self.route_preview_var, style="CardBody.TLabel", wraplength=980, justify="left").grid(
            row=3, column=0, columnspan=3, sticky="ew", pady=(8, 0)
        )

        center = ttk.Frame(page)
        center.grid(row=1, column=0, sticky="nsew", pady=(14, 14))
        center.grid_columnconfigure(0, weight=3)
        center.grid_columnconfigure(1, weight=2)
        center.grid_rowconfigure(0, weight=1)

        inputs = ttk.LabelFrame(center, text="Selected input files")
        inputs.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        inputs.grid_columnconfigure(0, weight=1)
        inputs.grid_rowconfigure(0, weight=1)
        inputs.grid_rowconfigure(3, weight=1)

        self.file_listbox = tk.Listbox(inputs, selectmode=tk.EXTENDED, relief="flat", borderwidth=1)
        self.file_listbox.grid(row=0, column=0, sticky="nsew")
        y_scroll = ttk.Scrollbar(inputs, orient="vertical", command=self.file_listbox.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(inputs, orient="horizontal", command=self.file_listbox.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.file_listbox.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        input_actions = ttk.Frame(inputs)
        input_actions.grid(row=0, column=2, sticky="ns", padx=(12, 0))
        ttk.Button(input_actions, text="Add files", command=self._add_files).grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(input_actions, text="Add folder", command=self._add_folder).grid(row=1, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(input_actions, text="Remove selected", command=self._remove_selected).grid(row=2, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(input_actions, text="Clear all", command=self._clear_inputs).grid(row=3, column=0, sticky="ew")

        ttk.Label(
            inputs,
            text="Choose a mode first. Folder scanning only pulls matching file extensions for that mode.",
            style="CardBody.TLabel",
            wraplength=700,
            justify="left",
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(10, 0))

        link_frame = ttk.LabelFrame(inputs, text="Online links / URLs")
        link_frame.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(14, 0))
        link_frame.grid_columnconfigure(0, weight=1)
        link_frame.grid_rowconfigure(2, weight=1)
        link_frame.grid_rowconfigure(4, weight=1)

        ttk.Label(
            link_frame,
            textvariable=self.link_status_summary_var,
            style="CardBody.TLabel",
            wraplength=700,
            justify="left",
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        link_actions = ttk.Frame(link_frame)
        link_actions.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 8))
        self.link_paste_button = ttk.Button(link_actions, text="Paste URLs", command=self._paste_urls_from_clipboard)
        self.link_paste_button.grid(row=0, column=0, padx=(0, 8))
        self.link_fetch_button = ttk.Button(link_actions, text="Fetch Links", command=self._start_link_fetch)
        self.link_fetch_button.grid(row=0, column=1, padx=(0, 8))
        self.link_fetch_start_button = ttk.Button(link_actions, text="Fetch + Start", command=lambda: self._start_link_fetch(auto_start=True))
        self.link_fetch_start_button.grid(row=0, column=2, padx=(0, 8))
        self.link_retry_button = ttk.Button(link_actions, text="Retry Failed", command=lambda: self._start_link_fetch(retry_failed=True))
        self.link_retry_button.grid(row=0, column=3, padx=(0, 8))
        self.link_pause_button = ttk.Button(link_actions, text="Pause", command=self._pause_link_fetch)
        self.link_pause_button.grid(row=0, column=4, padx=(0, 8))
        self.link_resume_button = ttk.Button(link_actions, text="Resume", command=self._resume_link_fetch)
        self.link_resume_button.grid(row=0, column=5, padx=(0, 8))
        self.link_cancel_button = ttk.Button(link_actions, text="Cancel", command=self._cancel_link_fetch)
        self.link_cancel_button.grid(row=0, column=6, padx=(0, 8))
        ttk.Button(link_actions, text="Clear URLs", command=self._clear_link_urls).grid(row=0, column=7, padx=(0, 8))
        ttk.Button(link_actions, text="Load Recent", command=self._load_recent_links).grid(row=0, column=8, padx=(0, 8))
        ttk.Button(link_actions, text="Open Cache", command=self._open_link_cache_dir).grid(row=0, column=9)

        self.link_input_text = ScrolledText(link_frame, wrap=tk.WORD, height=5, relief="flat", borderwidth=1)
        self.link_input_text.grid(row=2, column=0, columnspan=2, sticky="nsew")
        self.link_input_text.configure(padx=10, pady=8)

        ttk.Label(link_frame, textvariable=self.link_recent_summary_var, style="CardBody.TLabel", wraplength=700, justify="left").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(8, 6)
        )

        self.link_status_tree = ttk.Treeview(
            link_frame,
            columns=("status", "url", "file", "detail"),
            show="headings",
            height=5,
        )
        for col, title, width in (
            ("status", "Status", 110),
            ("url", "URL", 270),
            ("file", "Local file", 220),
            ("detail", "Detail", 120),
        ):
            self.link_status_tree.heading(col, text=title)
            self.link_status_tree.column(col, width=width, anchor="w")
        self.link_status_tree.grid(row=4, column=0, sticky="nsew")
        self.link_status_tree.bind("<<TreeviewSelect>>", self._on_link_status_selected)
        link_status_scroll = ttk.Scrollbar(link_frame, orient="vertical", command=self.link_status_tree.yview)
        link_status_scroll.grid(row=4, column=1, sticky="ns")
        self.link_status_tree.configure(yscrollcommand=link_status_scroll.set)

        ttk.Label(link_frame, textvariable=self.link_fetch_count_var, style="CardBody.TLabel").grid(row=5, column=0, sticky="w", pady=(8, 0))
        self.after_idle(self._update_link_fetch_buttons)

        options = ttk.LabelFrame(center, text="Output, options, and quick controls")
        options.grid(row=0, column=1, sticky="nsew")
        options.grid_columnconfigure(1, weight=1)
        options.grid_columnconfigure(2, weight=1)

        ttk.Label(options, text="Output folder:", style="Surface.TLabel").grid(row=0, column=0, sticky="w")
        self.output_entry = ttk.Entry(options, textvariable=self.output_dir_var)
        self.output_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=(8, 8))
        ttk.Button(options, text="Browse", command=self._browse_output_dir).grid(row=0, column=3, sticky="e")

        ttk.Button(options, text="Open folder", command=self._open_output_folder).grid(row=1, column=3, sticky="e", pady=(8, 0))

        self.merge_check = ttk.Checkbutton(options, text="Merge into 1 PDF", variable=self.merge_var, command=self._update_mode_controls)
        self.merge_check.grid(row=1, column=0, sticky="w", pady=(10, 0))

        ttk.Label(options, text="Merged output name:", style="Surface.TLabel").grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.output_name_entry = ttk.Entry(options, textvariable=self.output_name_var)
        self.output_name_entry.grid(row=2, column=1, columnspan=2, sticky="ew", padx=(8, 8), pady=(10, 0))

        self.recursive_check = ttk.Checkbutton(options, text="Scan folders recursively", variable=self.recursive_var)
        self.recursive_check.grid(row=3, column=0, sticky="w", pady=(10, 0))

        ttk.Label(options, text="Image format:", style="Surface.TLabel").grid(row=4, column=0, sticky="w", pady=(10, 0))
        self.image_format_combo = ttk.Combobox(options, textvariable=self.image_format_var, values=["png", "jpg"], state="readonly", width=10)
        self.image_format_combo.grid(row=4, column=1, sticky="w", padx=(8, 8), pady=(10, 0))

        ttk.Label(options, text="Image scale:", style="Surface.TLabel").grid(row=4, column=2, sticky="e", pady=(10, 0))
        self.image_scale_entry = ttk.Entry(options, textvariable=self.image_scale_var, width=10)
        self.image_scale_entry.grid(row=4, column=3, sticky="w", pady=(10, 0))

        ttk.Label(
            options,
            text="Scale 2.0 is a good default for PDF to image exports. Higher values increase size and clarity.",
            style="CardBody.TLabel",
            wraplength=360,
            justify="left",
        ).grid(row=5, column=0, columnspan=4, sticky="w", pady=(10, 0))

        ttk.Label(options, text="Dependency summary:", style="Surface.TLabel").grid(row=6, column=0, sticky="w", pady=(16, 0))
        ttk.Label(options, textvariable=self.dependency_var, style="CardBody.TLabel", wraplength=360, justify="left").grid(
            row=7, column=0, columnspan=4, sticky="w", pady=(6, 0)
        )

        log_frame = ttk.LabelFrame(page, text="Activity log and progress")
        log_frame.grid(row=2, column=0, sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)

        top_actions = ttk.Frame(log_frame)
        top_actions.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        top_actions.grid_columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(top_actions, mode="determinate", maximum=100)
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        self.start_button = ttk.Button(top_actions, text="Start conversion", style="Primary.TButton", command=self._start_conversion)
        self.start_button.grid(row=0, column=1, sticky="e")

        self.log_text = ScrolledText(log_frame, wrap=tk.WORD, height=12, relief="flat", borderwidth=1)
        self.log_text.grid(row=1, column=0, sticky="nsew")
        self.log_text.configure(state="disabled", padx=12, pady=10)

    def _build_pdf_tools_page(self) -> None:
        page = ttk.Frame(self.content)
        page.grid(row=0, column=0, sticky="nsew")
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(2, weight=1)
        self.pages["pdf_tools"] = page

        tool_card = ttk.LabelFrame(page, text="PDF studio: organize, secure, compress, redact, sign, and edit")
        tool_card.grid(row=0, column=0, sticky="ew")
        tool_card.grid_columnconfigure(1, weight=1)

        ttk.Label(tool_card, text="PDF tool:", style="Surface.TLabel").grid(row=0, column=0, sticky="w")
        self.pdf_tool_combo = ttk.Combobox(tool_card, textvariable=self.pdf_tool_var, values=PDF_TOOL_ORDER, state="readonly", width=34)
        self.pdf_tool_combo.grid(row=0, column=1, sticky="ew", padx=(10, 10))
        self.pdf_tool_combo.bind("<<ComboboxSelected>>", self._on_pdf_tool_changed)
        ttk.Button(tool_card, text="Open output folder", command=self._open_output_folder).grid(row=0, column=2, sticky="e")

        ttk.Label(tool_card, textvariable=self.pdf_tool_help_var, style="CardBody.TLabel", wraplength=980, justify="left").grid(
            row=1, column=0, columnspan=3, sticky="ew", pady=(10, 6)
        )
        ttk.Label(tool_card, textvariable=self.pdf_tool_hint_var, style="CardBody.TLabel", wraplength=980, justify="left").grid(
            row=2, column=0, columnspan=3, sticky="ew"
        )

        center = ttk.Frame(page)
        center.grid(row=1, column=0, sticky="nsew", pady=(14, 14))
        center.grid_columnconfigure(0, weight=3)
        center.grid_columnconfigure(1, weight=2)
        center.grid_rowconfigure(0, weight=1)

        inputs = ttk.LabelFrame(center, text="Selected PDF files")
        inputs.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        inputs.grid_columnconfigure(0, weight=1)
        inputs.grid_rowconfigure(0, weight=1)

        self.pdf_tool_listbox = tk.Listbox(inputs, selectmode=tk.EXTENDED, relief="flat", borderwidth=1)
        self.pdf_tool_listbox.grid(row=0, column=0, sticky="nsew")
        pdf_y_scroll = ttk.Scrollbar(inputs, orient="vertical", command=self.pdf_tool_listbox.yview)
        pdf_y_scroll.grid(row=0, column=1, sticky="ns")
        pdf_x_scroll = ttk.Scrollbar(inputs, orient="horizontal", command=self.pdf_tool_listbox.xview)
        pdf_x_scroll.grid(row=1, column=0, sticky="ew")
        self.pdf_tool_listbox.configure(yscrollcommand=pdf_y_scroll.set, xscrollcommand=pdf_x_scroll.set)

        input_actions = ttk.Frame(inputs)
        input_actions.grid(row=0, column=2, sticky="ns", padx=(12, 0))
        ttk.Button(input_actions, text="Add PDFs", command=self._add_pdf_tool_files).grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(input_actions, text="Add folder", command=self._add_pdf_tool_folder).grid(row=1, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(input_actions, text="Move up", command=self._move_pdf_tool_selected_up).grid(row=2, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(input_actions, text="Move down", command=self._move_pdf_tool_selected_down).grid(row=3, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(input_actions, text="Remove selected", command=self._remove_pdf_tool_selected).grid(row=4, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(input_actions, text="Clear all", command=self._clear_pdf_tool_inputs).grid(row=5, column=0, sticky="ew")

        ttk.Label(
            inputs,
            text="For merge, batch watermarking, password, and batch editing flows, the current list order is used. Use Move up and Move down when ordering matters.",
            style="CardBody.TLabel",
            wraplength=700,
            justify="left",
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(10, 0))

        options = ttk.LabelFrame(center, text="Tool options")
        options.grid(row=0, column=1, sticky="nsew")
        options.grid_columnconfigure(1, weight=1)
        options.grid_columnconfigure(3, weight=1)

        ttk.Label(options, text="Output folder:", style="Surface.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(options, textvariable=self.output_dir_var).grid(row=0, column=1, columnspan=2, sticky="ew", padx=(8, 8))
        ttk.Button(options, text="Browse", command=self._browse_output_dir).grid(row=0, column=3, sticky="e")

        ttk.Label(options, text="Merged output name:", style="Surface.TLabel").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.pdf_tool_output_name_entry = ttk.Entry(options, textvariable=self.pdf_tool_output_name_var)
        self.pdf_tool_output_name_entry.grid(row=1, column=1, columnspan=3, sticky="ew", padx=(8, 0), pady=(10, 0))

        self.pdf_tool_page_spec_label = ttk.Label(options, text="Pages / ranges:", style="Surface.TLabel")
        self.pdf_tool_page_spec_label.grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.pdf_tool_page_spec_entry = ttk.Entry(options, textvariable=self.pdf_tool_page_spec_var)
        self.pdf_tool_page_spec_entry.grid(row=2, column=1, columnspan=3, sticky="ew", padx=(8, 0), pady=(10, 0))

        ttk.Label(options, text="Split every N pages:", style="Surface.TLabel").grid(row=3, column=0, sticky="w", pady=(10, 0))
        self.pdf_tool_every_n_entry = ttk.Entry(options, textvariable=self.pdf_tool_every_n_var, width=10)
        self.pdf_tool_every_n_entry.grid(row=3, column=1, sticky="w", padx=(8, 0), pady=(10, 0))

        self.pdf_tool_text_label = ttk.Label(options, text="Watermark text:", style="Surface.TLabel")
        self.pdf_tool_text_label.grid(row=4, column=0, sticky="w", pady=(10, 0))
        self.pdf_tool_watermark_text_entry = ttk.Entry(options, textvariable=self.pdf_tool_watermark_text_var)
        self.pdf_tool_watermark_text_entry.grid(row=4, column=1, columnspan=3, sticky="ew", padx=(8, 0), pady=(10, 0))

        self.pdf_tool_image_label = ttk.Label(options, text="Watermark image:", style="Surface.TLabel")
        self.pdf_tool_image_label.grid(row=5, column=0, sticky="w", pady=(10, 0))
        self.pdf_tool_image_entry = ttk.Entry(options, textvariable=self.pdf_tool_watermark_image_var)
        self.pdf_tool_image_entry.grid(row=5, column=1, columnspan=2, sticky="ew", padx=(8, 8), pady=(10, 0))
        self.pdf_tool_image_browse_button = ttk.Button(options, text="Browse", command=self._browse_pdf_tool_watermark_image)
        self.pdf_tool_image_browse_button.grid(row=5, column=3, sticky="e", pady=(10, 0))

        self.pdf_tool_position_label = ttk.Label(options, text="Position:", style="Surface.TLabel")
        self.pdf_tool_position_label.grid(row=6, column=0, sticky="w", pady=(10, 0))
        self.pdf_tool_position_combo = ttk.Combobox(options, textvariable=self.pdf_tool_position_var, values=PDF_TOOL_POSITIONS, state="readonly", width=16)
        self.pdf_tool_position_combo.grid(row=6, column=1, sticky="w", padx=(8, 8), pady=(10, 0))

        self.pdf_tool_font_size_label = ttk.Label(options, text="Font size:", style="Surface.TLabel")
        self.pdf_tool_font_size_label.grid(row=6, column=2, sticky="e", pady=(10, 0))
        self.pdf_tool_font_size_entry = ttk.Entry(options, textvariable=self.pdf_tool_font_size_var, width=10)
        self.pdf_tool_font_size_entry.grid(row=6, column=3, sticky="w", pady=(10, 0))

        self.pdf_tool_rotation_label = ttk.Label(options, text="Rotation:", style="Surface.TLabel")
        self.pdf_tool_rotation_label.grid(row=7, column=0, sticky="w", pady=(10, 0))
        self.pdf_tool_rotation_entry = ttk.Entry(options, textvariable=self.pdf_tool_rotation_var, width=10)
        self.pdf_tool_rotation_entry.grid(row=7, column=1, sticky="w", padx=(8, 8), pady=(10, 0))

        self.pdf_tool_opacity_label = ttk.Label(options, text="Opacity:", style="Surface.TLabel")
        self.pdf_tool_opacity_label.grid(row=7, column=2, sticky="e", pady=(10, 0))
        self.pdf_tool_opacity_entry = ttk.Entry(options, textvariable=self.pdf_tool_opacity_var, width=10)
        self.pdf_tool_opacity_entry.grid(row=7, column=3, sticky="w", pady=(10, 0))

        self.pdf_tool_image_scale_label = ttk.Label(options, text="Image scale %:", style="Surface.TLabel")
        self.pdf_tool_image_scale_label.grid(row=8, column=0, sticky="w", pady=(10, 0))
        self.pdf_tool_image_scale_entry = ttk.Entry(options, textvariable=self.pdf_tool_image_scale_var, width=10)
        self.pdf_tool_image_scale_entry.grid(row=8, column=1, sticky="w", padx=(8, 0), pady=(10, 0))

        self.pdf_tool_metadata_title_label = ttk.Label(options, text="Metadata title:", style="Surface.TLabel")
        self.pdf_tool_metadata_title_label.grid(row=9, column=0, sticky="w", pady=(12, 0))
        self.pdf_tool_metadata_title_entry = ttk.Entry(options, textvariable=self.pdf_tool_metadata_title_var)
        self.pdf_tool_metadata_title_entry.grid(row=9, column=1, columnspan=3, sticky="ew", padx=(8, 0), pady=(12, 0))

        self.pdf_tool_metadata_author_label = ttk.Label(options, text="Metadata author:", style="Surface.TLabel")
        self.pdf_tool_metadata_author_label.grid(row=10, column=0, sticky="w", pady=(10, 0))
        self.pdf_tool_metadata_author_entry = ttk.Entry(options, textvariable=self.pdf_tool_metadata_author_var)
        self.pdf_tool_metadata_author_entry.grid(row=10, column=1, columnspan=3, sticky="ew", padx=(8, 0), pady=(10, 0))

        self.pdf_tool_metadata_subject_label = ttk.Label(options, text="Metadata subject:", style="Surface.TLabel")
        self.pdf_tool_metadata_subject_label.grid(row=11, column=0, sticky="w", pady=(10, 0))
        self.pdf_tool_metadata_subject_entry = ttk.Entry(options, textvariable=self.pdf_tool_metadata_subject_var)
        self.pdf_tool_metadata_subject_entry.grid(row=11, column=1, columnspan=3, sticky="ew", padx=(8, 0), pady=(10, 0))

        self.pdf_tool_metadata_keywords_label = ttk.Label(options, text="Metadata keywords:", style="Surface.TLabel")
        self.pdf_tool_metadata_keywords_label.grid(row=12, column=0, sticky="w", pady=(10, 0))
        self.pdf_tool_metadata_keywords_entry = ttk.Entry(options, textvariable=self.pdf_tool_metadata_keywords_var)
        self.pdf_tool_metadata_keywords_entry.grid(row=12, column=1, columnspan=3, sticky="ew", padx=(8, 0), pady=(10, 0))

        self.pdf_tool_metadata_clear_check = ttk.Checkbutton(
            options,
            text="Clear existing metadata before applying new values",
            variable=self.pdf_tool_metadata_clear_var,
        )
        self.pdf_tool_metadata_clear_check.grid(row=13, column=0, columnspan=4, sticky="w", pady=(10, 0))

        self.pdf_tool_redact_rect_label = ttk.Label(options, text="Area rectangle:", style="Surface.TLabel")
        self.pdf_tool_redact_rect_label.grid(row=14, column=0, sticky="w", pady=(10, 0))
        self.pdf_tool_redact_rect_entry = ttk.Entry(options, textvariable=self.pdf_tool_redact_rect_var)
        self.pdf_tool_redact_rect_entry.grid(row=14, column=1, columnspan=3, sticky="ew", padx=(8, 0), pady=(10, 0))

        self.pdf_tool_replacement_text_label = ttk.Label(options, text="Replacement text:", style="Surface.TLabel")
        self.pdf_tool_replacement_text_label.grid(row=15, column=0, sticky="w", pady=(10, 0))
        self.pdf_tool_replacement_text_entry = ttk.Entry(options, textvariable=self.pdf_tool_replacement_text_var)
        self.pdf_tool_replacement_text_entry.grid(row=15, column=1, columnspan=3, sticky="ew", padx=(8, 0), pady=(10, 0))

        self.pdf_tool_password_label = ttk.Label(options, text="Password:", style="Surface.TLabel")
        self.pdf_tool_password_label.grid(row=16, column=0, sticky="w", pady=(10, 0))
        self.pdf_tool_password_entry = ttk.Entry(options, textvariable=self.pdf_tool_password_var, show="•")
        self.pdf_tool_password_entry.grid(row=16, column=1, columnspan=3, sticky="ew", padx=(8, 0), pady=(10, 0))

        self.pdf_tool_owner_password_label = ttk.Label(options, text="Owner password:", style="Surface.TLabel")
        self.pdf_tool_owner_password_label.grid(row=17, column=0, sticky="w", pady=(10, 0))
        self.pdf_tool_owner_password_entry = ttk.Entry(options, textvariable=self.pdf_tool_owner_password_var, show="•")
        self.pdf_tool_owner_password_entry.grid(row=17, column=1, columnspan=3, sticky="ew", padx=(8, 0), pady=(10, 0))

        self.pdf_tool_compression_label = ttk.Label(options, text="Compression profile:", style="Surface.TLabel")
        self.pdf_tool_compression_label.grid(row=18, column=0, sticky="w", pady=(10, 0))
        self.pdf_tool_compression_combo = ttk.Combobox(
            options,
            textvariable=self.pdf_tool_compression_profile_var,
            values=["safe", "balanced", "strong"],
            state="readonly",
            width=16,
        )
        self.pdf_tool_compression_combo.grid(row=18, column=1, sticky="w", padx=(8, 0), pady=(10, 0))

        self.pdf_tool_options_help_label = ttk.Label(
            options,
            text="Examples: split ranges -> 1-3; 4-6; 8. Extract/remove -> 1-2, 5, 9-last. Reorder -> 3,1,2,2. Opacity is from 0.01 to 1.0.",
            style="CardBody.TLabel",
            wraplength=380,
            justify="left",
        )
        self.pdf_tool_options_help_label.grid(row=19, column=0, columnspan=4, sticky="w", pady=(12, 0))

        tool_log_frame = ttk.LabelFrame(page, text="PDF tool activity and progress")
        tool_log_frame.grid(row=2, column=0, sticky="nsew")
        tool_log_frame.grid_columnconfigure(0, weight=1)
        tool_log_frame.grid_rowconfigure(1, weight=1)

        top_actions = ttk.Frame(tool_log_frame)
        top_actions.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        top_actions.grid_columnconfigure(0, weight=1)
        self.pdf_tool_progress = ttk.Progressbar(top_actions, mode="determinate", maximum=100)
        self.pdf_tool_progress.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        ttk.Button(top_actions, text="Clear tool log", command=self._clear_pdf_tool_log).grid(row=0, column=1, sticky="e", padx=(0, 8))
        self.pdf_tool_start_button = ttk.Button(top_actions, text="Run PDF tool", style="Primary.TButton", command=self._start_pdf_tool)
        self.pdf_tool_start_button.grid(row=0, column=2, sticky="e")

        self.pdf_tool_log_text = ScrolledText(tool_log_frame, wrap=tk.WORD, height=12, relief="flat", borderwidth=1)
        self.pdf_tool_log_text.grid(row=1, column=0, sticky="nsew")
        self.pdf_tool_log_text.configure(state="disabled", padx=12, pady=10)

    def _build_ocr_page(self) -> None:
        page = ttk.Frame(self.content)
        page.grid(row=0, column=0, sticky="nsew")
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(2, weight=1)
        self.pages["ocr"] = page

        hero = ttk.Frame(page, style="Card.TFrame", padding=22)
        hero.grid(row=0, column=0, sticky="ew")
        hero.grid_columnconfigure(0, weight=1)
        ttk.Label(hero, text="Searchable PDF and OCR text", style="HeroTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            hero,
            text=(
                "Turn scans and screenshots into searchable PDFs or plain text. "
                "The OCR workspace supports searchable PDF generation, text extraction, saved OCR defaults, and progress updates that land in history like the rest of the app."
            ),
            style="CardBody.TLabel",
            wraplength=980,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))
        hero_actions = ttk.Frame(hero)
        hero_actions.grid(row=0, column=1, rowspan=2, sticky="e")
        ttk.Button(hero_actions, text="Settings", command=lambda: self._show_page("settings")).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(hero_actions, text="Test Tesseract", command=self._test_tesseract_path).grid(row=0, column=1)

        content = ttk.Frame(page)
        content.grid(row=1, column=0, sticky="nsew", pady=(14, 14))
        content.grid_columnconfigure(0, weight=3)
        content.grid_columnconfigure(1, weight=2)
        content.grid_rowconfigure(0, weight=1)

        inputs = ttk.LabelFrame(content, text="OCR source")
        inputs.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        inputs.grid_columnconfigure(1, weight=1)

        ttk.Label(inputs, text="Input image or PDF:", style="Surface.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(inputs, textvariable=self.ocr_input_var).grid(row=0, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(inputs, text="Browse", command=self._browse_ocr_input).grid(row=0, column=2, sticky="e")

        ttk.Label(inputs, text="Output folder:", style="Surface.TLabel").grid(row=1, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(inputs, textvariable=self.ocr_output_var).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=(10, 0))
        output_actions = ttk.Frame(inputs)
        output_actions.grid(row=1, column=2, sticky="e", pady=(10, 0))
        ttk.Button(output_actions, text="Browse", command=self._browse_ocr_output).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(output_actions, text="Open", command=lambda: open_path(Path(self.ocr_output_var.get().strip() or Path.cwd()))).grid(row=0, column=1)

        ttk.Label(
            inputs,
            text=(
                "Use Image -> Searchable PDF for single images or scans. Use PDF -> Searchable PDF to rasterize each page and add an invisible OCR text layer. "
                "Extract OCR Text saves a UTF-8 text file for review or downstream processing."
            ),
            style="CardBody.TLabel",
            wraplength=700,
            justify="left",
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(12, 0))

        settings = ttk.LabelFrame(content, text="OCR settings and dependency status")
        settings.grid(row=0, column=1, sticky="nsew")
        settings.grid_columnconfigure(1, weight=1)

        ttk.Label(settings, text="Language:", style="Surface.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(settings, textvariable=self.ocr_language_var, width=14).grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Label(settings, text="DPI:", style="Surface.TLabel").grid(row=1, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(settings, textvariable=self.ocr_dpi_var, width=14).grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(10, 0))
        ttk.Label(settings, text="PSM:", style="Surface.TLabel").grid(row=2, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(settings, textvariable=self.ocr_psm_var, width=14).grid(row=2, column=1, sticky="w", padx=(8, 0), pady=(10, 0))

        ttk.Label(settings, text="Tesseract path:", style="Surface.TLabel").grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(settings, textvariable=self.tesseract_path_var).grid(row=3, column=1, sticky="ew", padx=(8, 8), pady=(10, 0))
        tesseract_actions = ttk.Frame(settings)
        tesseract_actions.grid(row=3, column=2, sticky="e", pady=(10, 0))
        ttk.Button(tesseract_actions, text="Browse", command=self._browse_tesseract_path).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(tesseract_actions, text="Test", command=self._test_tesseract_path).grid(row=0, column=1)

        ttk.Label(settings, textvariable=self.ocr_dependency_var, style="CardBody.TLabel", wraplength=360, justify="left").grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(12, 0)
        )
        ttk.Label(settings, textvariable=self.ocr_status_var, style="CardBody.TLabel", wraplength=360, justify="left").grid(
            row=5, column=0, columnspan=3, sticky="w", pady=(8, 0)
        )

        action_row = ttk.Frame(settings)
        action_row.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(16, 0))
        for index in range(3):
            action_row.grid_columnconfigure(index, weight=1)
        self.ocr_start_image_button = ttk.Button(action_row, text="Image -> Searchable PDF", style="Primary.TButton", command=self._start_ocr_image_pdf)
        self.ocr_start_image_button.grid(row=0, column=0, sticky="ew")
        self.ocr_start_pdf_button = ttk.Button(action_row, text="PDF -> Searchable PDF", command=self._start_ocr_pdf_pdf)
        self.ocr_start_pdf_button.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self.ocr_start_text_button = ttk.Button(action_row, text="Extract OCR Text", command=self._start_ocr_text)
        self.ocr_start_text_button.grid(row=0, column=2, sticky="ew", padx=(8, 0))

        log_frame = ttk.LabelFrame(page, text="OCR log and progress")
        log_frame.grid(row=2, column=0, sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)

        top_actions = ttk.Frame(log_frame)
        top_actions.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        top_actions.grid_columnconfigure(0, weight=1)
        self.ocr_progress = ttk.Progressbar(top_actions, mode="determinate", maximum=100)
        self.ocr_progress.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        ttk.Button(top_actions, text="Clear OCR log", command=self._clear_ocr_log).grid(row=0, column=1, sticky="e")

        self.ocr_log_text = ScrolledText(log_frame, wrap=tk.WORD, height=12, relief="flat", borderwidth=1)
        self.ocr_log_text.grid(row=1, column=0, sticky="nsew")
        self.ocr_log_text.configure(state="disabled", padx=12, pady=10)

    def _build_pdf_tool_hint(self, tool: str) -> str:
        hints = {
            PDF_TOOL_MERGE: "Merge uses the visible file list order. Move files up or down before you run the tool.",
            PDF_TOOL_SPLIT_RANGES: "Use semicolons to mark each new output PDF. Example: 1-3; 4-6; 7,9-last.",
            PDF_TOOL_SPLIT_EVERY_N: "Enter a positive whole number. For example, 5 creates parts of 5 pages each.",
            PDF_TOOL_EXTRACT_PAGES: "Use commas for a page list. Example: 1-2, 5, 9-last.",
            PDF_TOOL_REMOVE_PAGES: "Use the same page syntax as extract. The selected pages will be dropped from the result.",
            PDF_TOOL_REORDER_PAGES: "Omitted pages will not appear in the result. You can also duplicate pages, such as 3,1,2,2.",
            PDF_TOOL_WATERMARK_TEXT: "Text watermark works on one or many PDFs. Try center + 45 degrees + opacity 0.18 for a classic diagonal stamp.",
            PDF_TOOL_WATERMARK_IMAGE: "PNG files with transparency usually work best. Scale is relative to page width.",
            PDF_TOOL_TEXT_OVERLAY: "Use this to add edit markers, review notes, labels, or approval text. Leave Pages / ranges blank to stamp all pages.",
            PDF_TOOL_IMAGE_OVERLAY: "Good for logos, approval stamps, badges, screenshots, or small diagrams placed on the page.",
            PDF_TOOL_REDACT_TEXT: "Search is case-insensitive. Blank Pages / ranges means every page. Matching text is permanently removed.",
            PDF_TOOL_REDACT_AREA: "Use x1,y1,x2,y2 in PDF points or percentages such as 10%,10%,90%,25%. Blank Pages / ranges targets all pages.",
            PDF_TOOL_EDIT_TEXT: "Best-effort replacement works only for extractable text. If nothing matches, use text overlay for a safe non-destructive note.",
            PDF_TOOL_SIGN_VISIBLE: "Blank Pages / ranges signs the last page. Add a signature image, typed signer text, or both.",
            PDF_TOOL_EDIT_METADATA: "Use Clear existing metadata to scrub current title, author, subject, keywords, and XML metadata before saving.",
            PDF_TOOL_LOCK: "Passwords are used only for the active run and are not written to recent history.",
            PDF_TOOL_UNLOCK: "Enter the current open password. The output copy is saved without encryption.",
            PDF_TOOL_COMPRESS: "Safe is fastest. Balanced is a strong default. Strong adds extra cleanup and duplicate-stream compression.",
        }
        return hints.get(tool, "")

    def _on_pdf_tool_changed(self, _event=None) -> None:
        self._update_pdf_tool_controls()

    def _update_pdf_tool_controls(self) -> None:
        tool = self.pdf_tool_var.get()
        self.pdf_tool_help_var.set(PDF_TOOL_HELP.get(tool, ""))
        self.pdf_tool_hint_var.set(self._build_pdf_tool_hint(tool))

        page_spec_enabled = tool in {
            PDF_TOOL_SPLIT_RANGES,
            PDF_TOOL_EXTRACT_PAGES,
            PDF_TOOL_REMOVE_PAGES,
            PDF_TOOL_REORDER_PAGES,
            PDF_TOOL_WATERMARK_TEXT,
            PDF_TOOL_WATERMARK_IMAGE,
            PDF_TOOL_TEXT_OVERLAY,
            PDF_TOOL_IMAGE_OVERLAY,
            PDF_TOOL_REDACT_TEXT,
            PDF_TOOL_REDACT_AREA,
            PDF_TOOL_EDIT_TEXT,
            PDF_TOOL_SIGN_VISIBLE,
        }
        every_n_enabled = tool == PDF_TOOL_SPLIT_EVERY_N
        merge_name_enabled = tool == PDF_TOOL_MERGE
        text_enabled = tool in {PDF_TOOL_WATERMARK_TEXT, PDF_TOOL_TEXT_OVERLAY, PDF_TOOL_REDACT_TEXT, PDF_TOOL_EDIT_TEXT, PDF_TOOL_SIGN_VISIBLE}
        image_enabled = tool in {PDF_TOOL_WATERMARK_IMAGE, PDF_TOOL_IMAGE_OVERLAY, PDF_TOOL_SIGN_VISIBLE}
        position_enabled = tool in {
            PDF_TOOL_WATERMARK_TEXT,
            PDF_TOOL_WATERMARK_IMAGE,
            PDF_TOOL_TEXT_OVERLAY,
            PDF_TOOL_IMAGE_OVERLAY,
            PDF_TOOL_SIGN_VISIBLE,
        }
        font_size_enabled = tool in {PDF_TOOL_WATERMARK_TEXT, PDF_TOOL_TEXT_OVERLAY}
        rotation_enabled = tool in {PDF_TOOL_WATERMARK_TEXT, PDF_TOOL_TEXT_OVERLAY}
        opacity_enabled = tool in {
            PDF_TOOL_WATERMARK_TEXT,
            PDF_TOOL_WATERMARK_IMAGE,
            PDF_TOOL_TEXT_OVERLAY,
            PDF_TOOL_IMAGE_OVERLAY,
            PDF_TOOL_SIGN_VISIBLE,
        }
        image_scale_enabled = tool in {PDF_TOOL_WATERMARK_IMAGE, PDF_TOOL_IMAGE_OVERLAY, PDF_TOOL_SIGN_VISIBLE}
        metadata_enabled = tool == PDF_TOOL_EDIT_METADATA
        rect_enabled = tool == PDF_TOOL_REDACT_AREA
        replacement_enabled = tool == PDF_TOOL_EDIT_TEXT
        password_enabled = tool in {PDF_TOOL_LOCK, PDF_TOOL_UNLOCK, PDF_TOOL_COMPRESS}
        owner_password_enabled = tool == PDF_TOOL_LOCK
        compression_enabled = tool == PDF_TOOL_COMPRESS

        page_label = "Pages / ranges:"
        if tool == PDF_TOOL_SIGN_VISIBLE:
            page_label = "Pages / ranges (blank = last page):"
        elif tool in {PDF_TOOL_REDACT_TEXT, PDF_TOOL_REDACT_AREA, PDF_TOOL_EDIT_TEXT}:
            page_label = "Pages / ranges (optional):"
        self.pdf_tool_page_spec_label.configure(text=page_label)

        text_label = "Text:"
        if tool == PDF_TOOL_WATERMARK_TEXT:
            text_label = "Watermark text:"
        elif tool == PDF_TOOL_TEXT_OVERLAY:
            text_label = "Overlay text:"
        elif tool == PDF_TOOL_REDACT_TEXT:
            text_label = "Search text to redact:"
        elif tool == PDF_TOOL_EDIT_TEXT:
            text_label = "Search text to replace:"
        elif tool == PDF_TOOL_SIGN_VISIBLE:
            text_label = "Signer text:"
        self.pdf_tool_text_label.configure(text=text_label)

        image_label = "Image file:"
        if tool == PDF_TOOL_WATERMARK_IMAGE:
            image_label = "Watermark image:"
        elif tool == PDF_TOOL_IMAGE_OVERLAY:
            image_label = "Overlay image:"
        elif tool == PDF_TOOL_SIGN_VISIBLE:
            image_label = "Signature image:"
        self.pdf_tool_image_label.configure(text=image_label)

        position_label = "Position:"
        if tool == PDF_TOOL_SIGN_VISIBLE:
            position_label = "Signature position:"
        self.pdf_tool_position_label.configure(text=position_label)
        self.pdf_tool_image_scale_label.configure(text="Signature scale %:" if tool == PDF_TOOL_SIGN_VISIBLE else "Image scale %:")
        self.pdf_tool_redact_rect_label.configure(text="Area rectangle:" if tool == PDF_TOOL_REDACT_AREA else "Area rectangle:")
        self.pdf_tool_replacement_text_label.configure(text="Replacement text:" if tool == PDF_TOOL_EDIT_TEXT else "Replacement text:")

        password_label = "Password:"
        if tool == PDF_TOOL_LOCK:
            password_label = "New user password:"
        elif tool == PDF_TOOL_UNLOCK:
            password_label = "Current password:"
        elif tool == PDF_TOOL_COMPRESS:
            password_label = "Password for encrypted PDF (optional):"
        self.pdf_tool_password_label.configure(text=password_label)
        self.pdf_tool_owner_password_label.configure(text="Owner password (optional):")

        if merge_name_enabled:
            self.pdf_tool_output_name_entry.state(["!disabled"])
        else:
            self.pdf_tool_output_name_entry.state(["disabled"])

        if page_spec_enabled:
            self.pdf_tool_page_spec_entry.state(["!disabled"])
        else:
            self.pdf_tool_page_spec_entry.state(["disabled"])

        if every_n_enabled:
            self.pdf_tool_every_n_entry.state(["!disabled"])
        else:
            self.pdf_tool_every_n_entry.state(["disabled"])

        if text_enabled:
            self.pdf_tool_watermark_text_entry.state(["!disabled"])
        else:
            self.pdf_tool_watermark_text_entry.state(["disabled"])

        if image_enabled:
            self.pdf_tool_image_entry.state(["!disabled"])
            self.pdf_tool_image_browse_button.state(["!disabled"])
        else:
            self.pdf_tool_image_entry.state(["disabled"])
            self.pdf_tool_image_browse_button.state(["disabled"])

        if position_enabled:
            self.pdf_tool_position_combo.state(["!disabled", "readonly"])
        else:
            self.pdf_tool_position_combo.state(["disabled"])

        if font_size_enabled:
            self.pdf_tool_font_size_entry.state(["!disabled"])
        else:
            self.pdf_tool_font_size_entry.state(["disabled"])

        if rotation_enabled:
            self.pdf_tool_rotation_entry.state(["!disabled"])
        else:
            self.pdf_tool_rotation_entry.state(["disabled"])

        if opacity_enabled:
            self.pdf_tool_opacity_entry.state(["!disabled"])
        else:
            self.pdf_tool_opacity_entry.state(["disabled"])

        if image_scale_enabled:
            self.pdf_tool_image_scale_entry.state(["!disabled"])
        else:
            self.pdf_tool_image_scale_entry.state(["disabled"])

        if rect_enabled:
            self.pdf_tool_redact_rect_entry.state(["!disabled"])
        else:
            self.pdf_tool_redact_rect_entry.state(["disabled"])

        if replacement_enabled:
            self.pdf_tool_replacement_text_entry.state(["!disabled"])
        else:
            self.pdf_tool_replacement_text_entry.state(["disabled"])

        for entry in (
            self.pdf_tool_metadata_title_entry,
            self.pdf_tool_metadata_author_entry,
            self.pdf_tool_metadata_subject_entry,
            self.pdf_tool_metadata_keywords_entry,
        ):
            if metadata_enabled:
                entry.state(["!disabled"])
            else:
                entry.state(["disabled"])

        if metadata_enabled:
            self.pdf_tool_metadata_clear_check.state(["!disabled"])
        else:
            self.pdf_tool_metadata_clear_check.state(["disabled"])

        if password_enabled:
            self.pdf_tool_password_entry.state(["!disabled"])
        else:
            self.pdf_tool_password_entry.state(["disabled"])

        if owner_password_enabled:
            self.pdf_tool_owner_password_entry.state(["!disabled"])
        else:
            self.pdf_tool_owner_password_entry.state(["disabled"])

        if compression_enabled:
            self.pdf_tool_compression_combo.state(["!disabled", "readonly"])
        else:
            self.pdf_tool_compression_combo.state(["disabled"])

        if tool == PDF_TOOL_EDIT_METADATA:
            self.pdf_tool_options_help_label.configure(
                text="Leave a metadata field blank to keep the current value unless Clear existing metadata is enabled."
            )
        elif tool == PDF_TOOL_SIGN_VISIBLE:
            self.pdf_tool_options_help_label.configure(
                text="Use bottom-right for a classic sign-off block. Opacity controls the full signature card transparency."
            )
        elif tool == PDF_TOOL_REDACT_TEXT:
            self.pdf_tool_options_help_label.configure(
                text="Redaction is permanent. Search is case-insensitive and all matches on the chosen pages are removed."
            )
        elif tool == PDF_TOOL_REDACT_AREA:
            self.pdf_tool_options_help_label.configure(
                text="Use points or percentages such as 36,72,420,160 or 10%,10%,90%,25%. Blank Pages / ranges means every page."
            )
        elif tool == PDF_TOOL_EDIT_TEXT:
            self.pdf_tool_options_help_label.configure(
                text="Best-effort replacement uses extractable text search plus redaction/replacement. For complex layouts, use text overlay instead."
            )
        elif tool == PDF_TOOL_LOCK:
            self.pdf_tool_options_help_label.configure(
                text="Passwords are kept only in memory for the active run and are never written to recent job history."
            )
        elif tool == PDF_TOOL_UNLOCK:
            self.pdf_tool_options_help_label.configure(
                text="Use the current open password. The output copy is saved without PDF encryption."
            )
        elif tool == PDF_TOOL_COMPRESS:
            self.pdf_tool_options_help_label.configure(
                text="Safe is fastest. Balanced is a strong default. Strong adds extra cleanup and duplicate-stream compression."
            )
        else:
            self.pdf_tool_options_help_label.configure(
                text="Examples: split ranges -> 1-3; 4-6; 8. Extract/remove -> 1-2, 5, 9-last. Reorder -> 3,1,2,2. Opacity is from 0.01 to 1.0."
            )

    def _add_pdf_tool_files(self) -> None:
        file_paths = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not file_paths:
            return
        self._append_pdf_tool_files([Path(path) for path in file_paths])
        self._show_page("pdf_tools")

    def _add_pdf_tool_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select PDF input folder")
        if not folder:
            return
        files = collect_files_from_folder(Path(folder), {".pdf"}, recursive=self.recursive_var.get())
        if not files:
            messagebox.showinfo("No PDFs found", "No PDF files were found in that folder.")
            return
        self._append_pdf_tool_files(files)
        self._show_page("pdf_tools")

    def _append_pdf_tool_files(self, new_files: list[Path]) -> None:
        merged: dict[str, Path] = {str(path.resolve()): path for path in self.pdf_tool_files}
        for path in new_files:
            merged.setdefault(str(path.resolve()), Path(path).resolve())
        self.pdf_tool_files = list(merged.values())
        self._refresh_pdf_tool_listbox()
        self.status_var.set(f"Loaded {len(self.pdf_tool_files)} PDF file(s) for PDF tools.")

    def _refresh_pdf_tool_listbox(self) -> None:
        self.pdf_tool_listbox.delete(0, tk.END)
        for path in self.pdf_tool_files:
            self.pdf_tool_listbox.insert(tk.END, str(path))

    def _move_pdf_tool_selected_up(self) -> None:
        indexes = list(self.pdf_tool_listbox.curselection())
        if not indexes or indexes[0] == 0:
            return
        for index in indexes:
            self.pdf_tool_files[index - 1], self.pdf_tool_files[index] = self.pdf_tool_files[index], self.pdf_tool_files[index - 1]
        self._refresh_pdf_tool_listbox()
        for index in [value - 1 for value in indexes]:
            self.pdf_tool_listbox.selection_set(index)
        self.status_var.set("Moved selected PDF item(s) up.")

    def _move_pdf_tool_selected_down(self) -> None:
        indexes = list(self.pdf_tool_listbox.curselection())
        if not indexes or indexes[-1] >= len(self.pdf_tool_files) - 1:
            return
        for index in reversed(indexes):
            self.pdf_tool_files[index], self.pdf_tool_files[index + 1] = self.pdf_tool_files[index + 1], self.pdf_tool_files[index]
        self._refresh_pdf_tool_listbox()
        for index in [value + 1 for value in indexes]:
            self.pdf_tool_listbox.selection_set(index)
        self.status_var.set("Moved selected PDF item(s) down.")

    def _remove_pdf_tool_selected(self) -> None:
        selected_indexes = list(self.pdf_tool_listbox.curselection())
        if not selected_indexes:
            return
        for index in reversed(selected_indexes):
            del self.pdf_tool_files[index]
        self._refresh_pdf_tool_listbox()
        self.status_var.set(f"Loaded {len(self.pdf_tool_files)} PDF file(s) for PDF tools.")

    def _clear_pdf_tool_inputs(self) -> None:
        self.pdf_tool_files.clear()
        self._refresh_pdf_tool_listbox()
        self.status_var.set("PDF tool inputs cleared.")

    def _browse_pdf_tool_watermark_image(self) -> None:
        image_path = filedialog.askopenfilename(
            title="Select image file",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.webp *.bmp *.tif *.tiff"), ("All files", "*.*")],
        )
        if image_path:
            self.pdf_tool_watermark_image_var.set(image_path)

    def _clear_pdf_tool_log(self) -> None:
        self.pdf_tool_log_text.configure(state="normal")
        self.pdf_tool_log_text.delete("1.0", tk.END)
        self.pdf_tool_log_text.configure(state="disabled")
        self.status_var.set("PDF tool log cleared.")

    def _append_pdf_tool_log(self, message: str) -> None:
        self.pdf_tool_log_text.configure(state="normal")
        self.pdf_tool_log_text.insert(tk.END, message + "\n")
        self.pdf_tool_log_text.see(tk.END)
        self.pdf_tool_log_text.configure(state="disabled")

    def _start_pdf_tool(self) -> None:
        if self.running:
            return
        if not self.pdf_tool_files:
            messagebox.showwarning("No PDF files selected", "Please add one or more PDF files first.")
            self._show_page("pdf_tools")
            return

        tool = self.pdf_tool_var.get()
        page_spec = self.pdf_tool_page_spec_var.get().strip()
        if tool in {PDF_TOOL_SPLIT_RANGES, PDF_TOOL_EXTRACT_PAGES, PDF_TOOL_REMOVE_PAGES, PDF_TOOL_REORDER_PAGES} and not page_spec:
            messagebox.showerror("Missing page specification", "Please enter the required pages or ranges for the selected PDF tool.")
            self._show_page("pdf_tools")
            return

        try:
            every_n_pages = int(self.pdf_tool_every_n_var.get().strip() or "1")
            font_size = int(self.pdf_tool_font_size_var.get().strip() or "42")
            rotation = float(self.pdf_tool_rotation_var.get().strip() or "45")
            opacity = float(self.pdf_tool_opacity_var.get().strip() or "0.18")
            image_scale = int(self.pdf_tool_image_scale_var.get().strip() or "40")
        except ValueError:
            messagebox.showerror(
                "Invalid PDF tool settings",
                "Please check numeric values such as split size, font size, rotation, opacity, and image scale.",
            )
            self._show_page("pdf_tools")
            return

        if tool == PDF_TOOL_SPLIT_EVERY_N and every_n_pages < 1:
            messagebox.showerror("Invalid split size", "Split every N pages must be a whole number greater than 0.")
            return

        text_value = self.pdf_tool_watermark_text_var.get().strip()
        watermark_image = self.pdf_tool_watermark_image_var.get().strip()
        password_value = self.pdf_tool_password_var.get()
        owner_password_value = self.pdf_tool_owner_password_var.get()
        compression_profile = self.pdf_tool_compression_profile_var.get().strip().lower() or "balanced"

        if tool in {PDF_TOOL_WATERMARK_TEXT, PDF_TOOL_TEXT_OVERLAY, PDF_TOOL_REDACT_TEXT, PDF_TOOL_EDIT_TEXT} and not text_value:
            messagebox.showerror("Missing text input", "Please enter the required text for the selected PDF tool.")
            return
        if tool in {PDF_TOOL_WATERMARK_IMAGE, PDF_TOOL_IMAGE_OVERLAY} and not watermark_image:
            messagebox.showerror("Missing image file", "Please choose an image file for the selected PDF tool.")
            return
        if tool == PDF_TOOL_SIGN_VISIBLE and not (text_value or watermark_image):
            messagebox.showerror("Missing signature input", "Please enter signer text, choose a signature image, or provide both.")
            return
        if tool == PDF_TOOL_REDACT_AREA and not self.pdf_tool_redact_rect_var.get().strip():
            messagebox.showerror("Missing area rectangle", "Please enter an area rectangle such as 36,72,420,160 or 10%,10%,90%,25%.")
            return
        if tool == PDF_TOOL_EDIT_TEXT and not self.pdf_tool_replacement_text_var.get().strip():
            messagebox.showerror("Missing replacement text", "Please enter replacement text for the best-effort edit tool.")
            return
        if tool == PDF_TOOL_EDIT_METADATA and not (
            self.pdf_tool_metadata_clear_var.get()
            or self.pdf_tool_metadata_title_var.get().strip()
            or self.pdf_tool_metadata_author_var.get().strip()
            or self.pdf_tool_metadata_subject_var.get().strip()
            or self.pdf_tool_metadata_keywords_var.get().strip()
        ):
            messagebox.showerror(
                "Missing metadata changes",
                "Enter at least one metadata field value or enable Clear existing metadata before running this tool.",
            )
            return
        if tool == PDF_TOOL_LOCK and not password_value:
            messagebox.showerror("Missing password", "Enter the new user password for the PDF lock tool.")
            return
        if tool == PDF_TOOL_UNLOCK and not password_value:
            messagebox.showerror("Missing password", "Enter the current password to unlock the selected PDF files.")
            return
        if tool == PDF_TOOL_COMPRESS and compression_profile not in {"safe", "balanced", "strong"}:
            messagebox.showerror("Invalid compression profile", "Compression profile must be safe, balanced, or strong.")
            return

        output_dir_text = self.output_dir_var.get().strip()
        if not output_dir_text:
            output_dir = self.pdf_tool_files[0].parent / "converted_output"
            self.output_dir_var.set(str(output_dir))
        else:
            output_dir = Path(output_dir_text).expanduser()

        config = PdfToolConfig(
            tool=tool,
            files=self.pdf_tool_files.copy(),
            output_dir=output_dir,
            output_name=self.pdf_tool_output_name_var.get().strip() or "merged_pdfs",
            page_spec=page_spec,
            every_n_pages=every_n_pages,
            watermark_text=text_value,
            watermark_image=Path(watermark_image).expanduser() if watermark_image else None,
            watermark_font_size=font_size,
            watermark_rotation=rotation,
            watermark_opacity=opacity,
            watermark_position=self.pdf_tool_position_var.get().strip() or "center",
            watermark_image_scale_percent=image_scale,
            metadata_title=self.pdf_tool_metadata_title_var.get().strip(),
            metadata_author=self.pdf_tool_metadata_author_var.get().strip(),
            metadata_subject=self.pdf_tool_metadata_subject_var.get().strip(),
            metadata_keywords=self.pdf_tool_metadata_keywords_var.get().strip(),
            metadata_clear_existing=bool(self.pdf_tool_metadata_clear_var.get()),
            redact_rect=self.pdf_tool_redact_rect_var.get().strip(),
            replacement_text=self.pdf_tool_replacement_text_var.get().strip(),
            pdf_password=password_value,
            pdf_owner_password=owner_password_value,
            compression_profile=compression_profile,
        )

        self.running = True
        self.active_run_kind = "pdf_tool"
        self._set_ocr_buttons_enabled(False)
        self.pdf_tool_progress["value"] = 0
        self.status_var.set("Running PDF tool...")
        self._append_pdf_tool_log("\n=== New PDF tool run started ===")
        self._append_pdf_tool_log(f"Tool: {config.tool}")
        self._append_pdf_tool_log(f"Output folder: {config.output_dir}")
        self._append_pdf_tool_log(f"PDFs selected: {len(config.files)}")
        if config.tool == PDF_TOOL_COMPRESS:
            self._append_pdf_tool_log(f"Compression profile: {config.compression_profile}")
        if config.tool in {PDF_TOOL_LOCK, PDF_TOOL_UNLOCK, PDF_TOOL_COMPRESS}:
            self._append_pdf_tool_log("Password data captured in memory for this run only.")

        worker = threading.Thread(target=self._pdf_tool_worker_run, args=(config,), daemon=True)
        worker.start()
        self.after(100, self._poll_worker_queue)
        self._show_page("pdf_tools")

    def _pdf_tool_worker_run(self, config: PdfToolConfig) -> None:
        def log(message: str) -> None:
            self.worker_queue.put(("pdf_log", message))

        def progress(current: int, total: int) -> None:
            self.worker_queue.put(("pdf_progress", (current, total)))

        try:
            outputs = process_pdf_tool(config, log=log, progress=progress)
            self.worker_queue.put(("pdf_done", (config, outputs)))
        except Exception:
            self.worker_queue.put(("pdf_error", (config, traceback.format_exc())))

    def _build_organizer_page(self) -> None:
        page = ttk.Frame(self.content)
        page.grid(row=0, column=0, sticky="nsew")
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(0, weight=1)
        self.pages["organizer"] = page

        self.organizer_panel = PageOrganizerPanel(
            page,
            palette=resolve_palette(self.theme_choice_var.get()),
            set_status=self.status_var.set,
            on_recent_job=self._record_external_job,
            on_loaded_pdf=self._remember_organizer_pdf,
        )
        self.organizer_panel.grid(row=0, column=0, sticky="nsew")

    def _open_pdf_in_organizer(self) -> None:
        self._show_page("organizer")
        if self.organizer_panel is not None:
            self.organizer_panel.open_pdf_dialog()

    def _record_external_job(self, record: dict[str, object]) -> None:
        self.state_store.add_recent_job(record)
        if str(status).lower() != "completed":
            self.state_store.add_failed_job(dict(record))
        self._refresh_history_views()

    def _track_success_outputs(self, outputs: list[str | Path], *, output_dir: Path | None = None, label: str = "") -> None:
        normalized = [Path(item) for item in outputs if str(item).strip()]
        self.last_outputs = normalized
        if output_dir is not None:
            self.last_output_dir = Path(output_dir)
        elif normalized:
            first = normalized[0]
            self.last_output_dir = first if first.is_dir() else first.parent
        self.last_job_label = label
        if normalized:
            self.state_store.remember_outputs(normalized)
        self._refresh_recent_outputs_view()

    def _maybe_auto_open_output(self, output_dir: Path | None) -> None:
        if not output_dir or not bool(self.auto_open_output_var.get()):
            return
        try:
            folder = Path(output_dir)
            folder.mkdir(parents=True, exist_ok=True)
            open_path(folder)
        except Exception:
            pass

    def _selected_recent_output_path(self) -> Path | None:
        if not hasattr(self, "recent_outputs_listbox"):
            return None
        selection = self.recent_outputs_listbox.curselection()
        if not selection:
            return None
        return Path(self.recent_output_item_ids.get(str(selection[0]), "")) if self.recent_output_item_ids.get(str(selection[0]), "") else None

    def _open_selected_recent_output(self) -> None:
        path = self._selected_recent_output_path()
        if not path:
            messagebox.showinfo("Recent outputs", "Select a recent output first.")
            return
        if not path.exists():
            messagebox.showwarning("Recent outputs", f"This output is no longer available:\n{path}")
            return
        open_path(path)

    def _open_recent_output_folder(self) -> None:
        path = self._selected_recent_output_path()
        if not path:
            messagebox.showinfo("Recent outputs", "Select a recent output first.")
            return
        target = path if path.is_dir() else path.parent
        if not target.exists():
            messagebox.showwarning("Recent outputs", f"This output folder is no longer available:\n{target}")
            return
        open_path(target)

    def _clear_recent_outputs(self) -> None:
        self.state_store.clear_recent_outputs()
        self._refresh_recent_outputs_view()
        self.status_var.set("Recent outputs cleared.")

    def _refresh_recent_outputs_view(self) -> None:
        if not hasattr(self, "recent_outputs_listbox"):
            return
        self.recent_outputs_listbox.delete(0, tk.END)
        self.recent_output_item_ids.clear()
        for index, item in enumerate(self.state_store.recent_outputs()):
            path = Path(item)
            display = path.name or item
            folder_name = path.parent.name if str(path.parent) not in {"", "."} else ""
            label = display if not folder_name else f"{display}  •  {folder_name}"
            self.recent_outputs_listbox.insert(tk.END, label)
            self.recent_output_item_ids[str(index)] = item

    def _selected_failed_job(self) -> dict[str, object] | None:
        if not hasattr(self, "failed_jobs_listbox"):
            return None
        selection = self.failed_jobs_listbox.curselection()
        if not selection:
            return None
        return self.failed_job_item_ids.get(str(selection[0]))

    def _remove_selected_failed_job(self) -> None:
        job = self._selected_failed_job()
        if not job:
            messagebox.showinfo("Failed jobs", "Select a failed job first.")
            return
        self.state_store.remove_failed_job(str(job.get("id", "")))
        self._refresh_failed_jobs_view()
        self.status_var.set("Removed the selected failed job.")

    def _clear_failed_jobs(self) -> None:
        self.state_store.clear_failed_jobs()
        self._refresh_failed_jobs_view()
        self.status_var.set("Failed jobs cleared.")

    def _refresh_failed_jobs_view(self) -> None:
        if not hasattr(self, "failed_jobs_listbox"):
            return
        self.failed_jobs_listbox.delete(0, tk.END)
        self.failed_job_item_ids.clear()
        for index, job in enumerate(self.state_store.failed_jobs()):
            mode = str(job.get("mode", job.get("tool", "Failed job"))).strip() or "Failed job"
            timestamp = str(job.get("timestamp", "")).strip()
            file_count = job.get("file_count", 0)
            label = f"{timestamp}  |  {mode}  |  files: {file_count}"
            self.failed_jobs_listbox.insert(tk.END, label)
            self.failed_job_item_ids[str(index)] = job

    def _retry_selected_failed_job(self) -> None:
        if self.running:
            return
        job = self._selected_failed_job()
        if not job:
            messagebox.showinfo("Failed jobs", "Select a failed job first.")
            return
        file_paths = [Path(item) for item in (job.get("input_files", []) or []) if str(item).strip()]
        existing = [path for path in file_paths if path.exists()]
        missing_count = len(file_paths) - len(existing)
        if not existing:
            messagebox.showerror("Failed jobs", "The original input files are missing, so this job cannot be retried.")
            return
        if missing_count:
            self.status_var.set(f"Retrying with {len(existing)} existing file(s). {missing_count} file(s) are missing.")
        job_type = str(job.get("job_type", "convert"))
        self.output_dir_var.set(str(job.get("output_dir", self.output_dir_var.get())))
        if job_type == "pdf_tool":
            self.pdf_tool_files = existing
            self._refresh_pdf_tool_listbox()
            self.pdf_tool_var.set(str(job.get("tool", PDF_TOOL_MERGE)))
            self.pdf_tool_output_name_var.set(str(job.get("output_name", "merged_pdfs")))
            self.pdf_tool_page_spec_var.set(str(job.get("page_spec", "")))
            self.pdf_tool_every_n_var.set(str(job.get("every_n_pages", "2")))
            self.pdf_tool_watermark_text_var.set(str(job.get("watermark_text", "")))
            self.pdf_tool_watermark_image_var.set(str(job.get("watermark_image", "")))
            self.pdf_tool_font_size_var.set(str(job.get("watermark_font_size", "42")))
            self.pdf_tool_rotation_var.set(str(job.get("watermark_rotation", "45")))
            self.pdf_tool_opacity_var.set(str(job.get("watermark_opacity", "0.18")))
            self.pdf_tool_position_var.set(str(job.get("watermark_position", "center")))
            self.pdf_tool_image_scale_var.set(str(job.get("watermark_image_scale_percent", "40")))
            self.pdf_tool_metadata_title_var.set(str(job.get("metadata_title", "")))
            self.pdf_tool_metadata_author_var.set(str(job.get("metadata_author", "")))
            self.pdf_tool_metadata_subject_var.set(str(job.get("metadata_subject", "")))
            self.pdf_tool_metadata_keywords_var.set(str(job.get("metadata_keywords", "")))
            self.pdf_tool_metadata_clear_var.set(bool(job.get("metadata_clear_existing", False)))
            self.pdf_tool_redact_rect_var.set(str(job.get("redact_rect", self.pdf_tool_redact_rect_var.get())))
            self.pdf_tool_replacement_text_var.set(str(job.get("replacement_text", "")))
            self.pdf_tool_compression_profile_var.set(str(job.get("compression_profile", "balanced")))
            self.pdf_tool_password_var.set("")
            self.pdf_tool_owner_password_var.set("")
            self._update_pdf_tool_controls()
            self._show_page("pdf_tools")
            self.after(60, self._start_pdf_tool)
            return
        if job_type == "ocr":
            self.ocr_input_var.set(str(existing[0]))
            self.ocr_output_var.set(str(job.get("output_dir", self.ocr_output_var.get())))
            self.ocr_language_var.set(str(job.get("ocr_language", self.ocr_language_var.get())))
            self.ocr_dpi_var.set(str(job.get("ocr_dpi", self.ocr_dpi_var.get())))
            self.ocr_psm_var.set(str(job.get("ocr_psm", self.ocr_psm_var.get())))
            self._show_page("ocr")
            mode = str(job.get("mode", "OCR"))
            if "Image -> Searchable PDF" in mode:
                self.after(60, self._start_ocr_image_pdf)
            elif "PDF -> Searchable PDF" in mode:
                self.after(60, self._start_ocr_pdf_pdf)
            else:
                self.after(60, self._start_ocr_text)
            return
        self.selected_files = existing
        self._refresh_file_listbox()
        self.mode_var.set(str(job.get("mode", MODE_ANY_TO_PDF)))
        self.merge_var.set(bool(job.get("merge_to_one_pdf", False)))
        self.output_name_var.set(str(job.get("merged_output_name", default_merged_name(self.mode_var.get()))))
        self.image_format_var.set(str(job.get("image_format", "png")))
        self.image_scale_var.set(str(job.get("image_scale", "2.0")))
        self.engine_mode_var.set(str(job.get("engine_mode", ENGINE_AUTO)))
        self._refresh_dependency_status()
        self._update_mode_controls()
        self._show_page("convert")
        self.after(60, self._start_conversion)

    def _capture_session_snapshot(self) -> dict[str, object]:
        link_text = ""
        if hasattr(self, "link_input_text"):
            try:
                link_text = self.link_input_text.get("1.0", tk.END).strip()
            except Exception:
                link_text = ""
        return {
            "selected_files": [str(path) for path in self.selected_files],
            "pdf_tool_files": [str(path) for path in self.pdf_tool_files],
            "mode": self.mode_var.get(),
            "output_dir": self.output_dir_var.get().strip(),
            "merge_to_one_pdf": bool(self.merge_var.get()),
            "merged_output_name": self.output_name_var.get().strip(),
            "image_format": self.image_format_var.get().strip(),
            "image_scale": self.image_scale_var.get().strip(),
            "engine_mode": self.engine_mode_var.get().strip(),
            "current_page": self.current_page,
            "link_input_text": link_text,
            "pdf_tool": self.pdf_tool_var.get(),
            "pdf_tool_output_name": self.pdf_tool_output_name_var.get().strip(),
            "pdf_tool_page_spec": self.pdf_tool_page_spec_var.get().strip(),
            "pdf_tool_every_n": self.pdf_tool_every_n_var.get().strip(),
            "pdf_tool_text": self.pdf_tool_watermark_text_var.get(),
            "pdf_tool_image": self.pdf_tool_watermark_image_var.get().strip(),
            "pdf_tool_font_size": self.pdf_tool_font_size_var.get().strip(),
            "pdf_tool_rotation": self.pdf_tool_rotation_var.get().strip(),
            "pdf_tool_opacity": self.pdf_tool_opacity_var.get().strip(),
            "pdf_tool_position": self.pdf_tool_position_var.get().strip(),
            "pdf_tool_image_scale": self.pdf_tool_image_scale_var.get().strip(),
            "pdf_tool_metadata_title": self.pdf_tool_metadata_title_var.get().strip(),
            "pdf_tool_metadata_author": self.pdf_tool_metadata_author_var.get().strip(),
            "pdf_tool_metadata_subject": self.pdf_tool_metadata_subject_var.get().strip(),
            "pdf_tool_metadata_keywords": self.pdf_tool_metadata_keywords_var.get().strip(),
            "pdf_tool_metadata_clear": bool(self.pdf_tool_metadata_clear_var.get()),
            "pdf_tool_redact_rect": self.pdf_tool_redact_rect_var.get().strip(),
            "pdf_tool_replacement_text": self.pdf_tool_replacement_text_var.get(),
            "pdf_tool_compression_profile": self.pdf_tool_compression_profile_var.get().strip(),
        }

    def _restore_last_session_snapshot(self) -> None:
        if not bool(self.restore_session_var.get()):
            return
        snapshot = self.state_store.session_snapshot()
        if not snapshot:
            return
        selected_files = [Path(item) for item in snapshot.get("selected_files", []) if Path(item).exists()]
        pdf_tool_files = [Path(item) for item in snapshot.get("pdf_tool_files", []) if Path(item).exists()]
        if selected_files:
            self.selected_files = selected_files
            self._refresh_file_listbox()
        if pdf_tool_files:
            self.pdf_tool_files = pdf_tool_files
            self._refresh_pdf_tool_listbox()
        self.mode_var.set(str(snapshot.get("mode", self.mode_var.get())))
        self.output_dir_var.set(str(snapshot.get("output_dir", self.output_dir_var.get())))
        self.merge_var.set(bool(snapshot.get("merge_to_one_pdf", self.merge_var.get())))
        self.output_name_var.set(str(snapshot.get("merged_output_name", self.output_name_var.get())))
        self.image_format_var.set(str(snapshot.get("image_format", self.image_format_var.get())))
        self.image_scale_var.set(str(snapshot.get("image_scale", self.image_scale_var.get())))
        self.engine_mode_var.set(str(snapshot.get("engine_mode", self.engine_mode_var.get())))
        self.pdf_tool_var.set(str(snapshot.get("pdf_tool", self.pdf_tool_var.get())))
        self.pdf_tool_output_name_var.set(str(snapshot.get("pdf_tool_output_name", self.pdf_tool_output_name_var.get())))
        self.pdf_tool_page_spec_var.set(str(snapshot.get("pdf_tool_page_spec", self.pdf_tool_page_spec_var.get())))
        self.pdf_tool_every_n_var.set(str(snapshot.get("pdf_tool_every_n", self.pdf_tool_every_n_var.get())))
        self.pdf_tool_watermark_text_var.set(str(snapshot.get("pdf_tool_text", self.pdf_tool_watermark_text_var.get())))
        self.pdf_tool_watermark_image_var.set(str(snapshot.get("pdf_tool_image", self.pdf_tool_watermark_image_var.get())))
        self.pdf_tool_font_size_var.set(str(snapshot.get("pdf_tool_font_size", self.pdf_tool_font_size_var.get())))
        self.pdf_tool_rotation_var.set(str(snapshot.get("pdf_tool_rotation", self.pdf_tool_rotation_var.get())))
        self.pdf_tool_opacity_var.set(str(snapshot.get("pdf_tool_opacity", self.pdf_tool_opacity_var.get())))
        self.pdf_tool_position_var.set(str(snapshot.get("pdf_tool_position", self.pdf_tool_position_var.get())))
        self.pdf_tool_image_scale_var.set(str(snapshot.get("pdf_tool_image_scale", self.pdf_tool_image_scale_var.get())))
        self.pdf_tool_metadata_title_var.set(str(snapshot.get("pdf_tool_metadata_title", self.pdf_tool_metadata_title_var.get())))
        self.pdf_tool_metadata_author_var.set(str(snapshot.get("pdf_tool_metadata_author", self.pdf_tool_metadata_author_var.get())))
        self.pdf_tool_metadata_subject_var.set(str(snapshot.get("pdf_tool_metadata_subject", self.pdf_tool_metadata_subject_var.get())))
        self.pdf_tool_metadata_keywords_var.set(str(snapshot.get("pdf_tool_metadata_keywords", self.pdf_tool_metadata_keywords_var.get())))
        self.pdf_tool_metadata_clear_var.set(bool(snapshot.get("pdf_tool_metadata_clear", self.pdf_tool_metadata_clear_var.get())))
        self.pdf_tool_redact_rect_var.set(str(snapshot.get("pdf_tool_redact_rect", self.pdf_tool_redact_rect_var.get())))
        self.pdf_tool_replacement_text_var.set(str(snapshot.get("pdf_tool_replacement_text", self.pdf_tool_replacement_text_var.get())))
        self.pdf_tool_compression_profile_var.set(str(snapshot.get("pdf_tool_compression_profile", self.pdf_tool_compression_profile_var.get())))
        if hasattr(self, "link_input_text"):
            try:
                self.link_input_text.delete("1.0", tk.END)
                link_text = str(snapshot.get("link_input_text", "")).strip()
                if link_text:
                    self.link_input_text.insert("1.0", link_text)
            except Exception:
                pass
        self._refresh_dependency_status()
        self._update_mode_controls()
        self._update_pdf_tool_controls()
        self._refresh_home_summary()
        target_page = str(snapshot.get("current_page", "home")).strip()
        if target_page in self.pages:
            self._show_page(target_page)
        restored_parts = []
        if selected_files:
            restored_parts.append(f"{len(selected_files)} convert input(s)")
        if pdf_tool_files:
            restored_parts.append(f"{len(pdf_tool_files)} PDF tool input(s)")
        if restored_parts:
            self.status_var.set("Restored last session: " + ", ".join(restored_parts))

    def _collect_log_sections(self) -> str:
        sections: list[str] = []
        for title, widget_name in (
            ("Main activity log", "log_text"),
            ("PDF tool log", "pdf_tool_log_text"),
            ("OCR log", "ocr_log_text"),
            ("Automation log", "automation_log_text"),
        ):
            widget = getattr(self, widget_name, None)
            content = ""
            if widget is not None:
                try:
                    content = widget.get("1.0", tk.END).strip()
                except Exception:
                    content = ""
            sections.append(f"===== {title} =====\n{content or '(empty)'}")
        return "\n\n".join(sections) + "\n"

    def _export_current_logs_action(self) -> None:
        default_dir = self.last_output_dir or Path(self.output_dir_var.get().strip() or Path.cwd())
        filename = f"gokul_omni_convert_logs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        target = filedialog.asksaveasfilename(
            title="Export app logs",
            defaultextension=".txt",
            initialdir=str(default_dir),
            initialfile=filename,
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if not target:
            return
        path = export_text_file(Path(target), self._collect_log_sections())
        self.status_var.set(f"Exported logs: {path}")
        open_path(path)

    def _check_for_updates_placeholder(self) -> None:
        now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.last_update_check_var.set(now_text)
        self.state_store.set("last_update_check", now_text)
        messagebox.showinfo(
            "Update checker",
            "Patch 15 still includes an update-check placeholder only.\n\nUse this hook later for GitHub releases, an installer feed, or your own version endpoint.",
        )
        self.status_var.set("Update check placeholder completed.")

    def _remember_organizer_pdf(self, path: Path) -> None:
        self.state_store.set("organizer_last_pdf", str(path))

    def _build_history_page(self) -> None:
        page = ttk.Frame(self.content)
        page.grid(row=0, column=0, sticky="nsew")
        page.grid_columnconfigure(0, weight=2)
        page.grid_columnconfigure(1, weight=1)
        page.grid_rowconfigure(0, weight=1)
        self.pages["history"] = page

        left = ttk.LabelFrame(page, text="Recent jobs")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(1, weight=1)

        ttk.Label(
            left,
            text="This view stores job summaries locally so you can inspect past runs and load the same settings again.",
            style="CardBody.TLabel",
            wraplength=720,
            justify="left",
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        self.history_tree = ttk.Treeview(
            left,
            columns=("time", "status", "mode", "files", "outputs"),
            show="headings",
            height=14,
        )
        for col, title, width in (
            ("time", "Time", 170),
            ("status", "Status", 120),
            ("mode", "Mode", 360),
            ("files", "Files", 70),
            ("outputs", "Outputs", 80),
        ):
            self.history_tree.heading(col, text=title)
            self.history_tree.column(col, width=width, anchor="w")
        self.history_tree.grid(row=1, column=0, sticky="nsew")
        self.history_tree.bind("<<TreeviewSelect>>", self._on_history_selected)

        right = ttk.LabelFrame(page, text="Selected job details")
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(1, weight=1)

        ttk.Label(right, textvariable=self.history_detail_var, style="CardBody.TLabel", wraplength=340, justify="left").grid(
            row=0, column=0, sticky="nw"
        )
        self.history_details_text = ScrolledText(right, wrap=tk.WORD, relief="flat", borderwidth=1, height=16, padx=12, pady=10)
        self.history_details_text.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        self.history_details_text.configure(state="disabled")

        action_row = ttk.Frame(right)
        action_row.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        action_row.grid_columnconfigure(0, weight=1)
        ttk.Button(action_row, text="Load selected settings", command=self._load_selected_history_settings).grid(row=0, column=0, sticky="ew")
        ttk.Button(action_row, text="Open selected output folder", command=self._open_selected_history_output).grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(action_row, text="Export logs", command=self._export_current_logs_action).grid(row=2, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(action_row, text="Clear history", command=self._clear_history).grid(row=3, column=0, sticky="ew", pady=(8, 0))

        right.grid_rowconfigure(3, weight=1)
        lower_panels = ttk.Frame(right)
        lower_panels.grid(row=3, column=0, sticky="nsew", pady=(12, 0))
        lower_panels.grid_columnconfigure(0, weight=1)
        lower_panels.grid_columnconfigure(1, weight=1)
        lower_panels.grid_rowconfigure(1, weight=1)

        ttk.Label(lower_panels, text="Recent outputs", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(lower_panels, text="Failed jobs ready for retry", style="CardTitle.TLabel").grid(row=0, column=1, sticky="w", padx=(10, 0))

        recent_frame = ttk.Frame(lower_panels)
        recent_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        recent_frame.grid_columnconfigure(0, weight=1)
        recent_frame.grid_rowconfigure(0, weight=1)
        self.recent_outputs_listbox = tk.Listbox(recent_frame, selectmode=tk.SINGLE, exportselection=False, relief="flat", borderwidth=1)
        self.recent_outputs_listbox.grid(row=0, column=0, sticky="nsew")
        recent_scroll = ttk.Scrollbar(recent_frame, orient="vertical", command=self.recent_outputs_listbox.yview)
        recent_scroll.grid(row=0, column=1, sticky="ns")
        self.recent_outputs_listbox.configure(yscrollcommand=recent_scroll.set)
        recent_actions = ttk.Frame(recent_frame)
        recent_actions.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        recent_actions.grid_columnconfigure(0, weight=1)
        ttk.Button(recent_actions, text="Open file", command=self._open_selected_recent_output).grid(row=0, column=0, sticky="ew")
        ttk.Button(recent_actions, text="Open folder", command=self._open_recent_output_folder).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        ttk.Button(recent_actions, text="Clear", command=self._clear_recent_outputs).grid(row=0, column=2, sticky="ew", padx=(8, 0))

        failed_frame = ttk.Frame(lower_panels)
        failed_frame.grid(row=1, column=1, sticky="nsew")
        failed_frame.grid_columnconfigure(0, weight=1)
        failed_frame.grid_rowconfigure(0, weight=1)
        self.failed_jobs_listbox = tk.Listbox(failed_frame, selectmode=tk.SINGLE, exportselection=False, relief="flat", borderwidth=1)
        self.failed_jobs_listbox.grid(row=0, column=0, sticky="nsew")
        failed_scroll = ttk.Scrollbar(failed_frame, orient="vertical", command=self.failed_jobs_listbox.yview)
        failed_scroll.grid(row=0, column=1, sticky="ns")
        self.failed_jobs_listbox.configure(yscrollcommand=failed_scroll.set)
        failed_actions = ttk.Frame(failed_frame)
        failed_actions.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        failed_actions.grid_columnconfigure(0, weight=1)
        ttk.Button(failed_actions, text="Retry selected", command=self._retry_selected_failed_job).grid(row=0, column=0, sticky="ew")
        ttk.Button(failed_actions, text="Remove selected", command=self._remove_selected_failed_job).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        ttk.Button(failed_actions, text="Clear all", command=self._clear_failed_jobs).grid(row=0, column=2, sticky="ew", padx=(8, 0))

    def _build_settings_page(self) -> None:
        page = ttk.Frame(self.content)
        page.grid(row=0, column=0, sticky="nsew")
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(1, weight=1)
        self.pages["settings"] = page

        header = ttk.Frame(page, style="Card.TFrame", padding=22)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        ttk.Label(header, text="Startup, engine selection, reminders, OCR, and local files", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text=(
                "Patch 15 adds workflow acceleration on top of the startup polish from earlier patches, including favorite presets, a quick-actions palette, link pause/resume, cache controls, and performance tuning while keeping pure Python as the default engine."
            ),
            style="CardBody.TLabel",
            wraplength=980,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        content = ttk.Frame(page)
        content.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        content.grid_columnconfigure(0, weight=1)
        content.grid_columnconfigure(1, weight=1)
        content.grid_rowconfigure(0, weight=1)
        content.grid_rowconfigure(1, weight=1)

        appearance = ttk.LabelFrame(content, text="Theme and behavior")
        appearance.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=(0, 10))
        appearance.grid_columnconfigure(1, weight=1)

        ttk.Label(appearance, text="Theme:", style="Surface.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Combobox(appearance, textvariable=self.theme_choice_var, values=["dark", "light", "system"], state="readonly", width=12).grid(
            row=0, column=1, sticky="w", padx=(8, 0)
        )
        ttk.Button(appearance, text="Apply theme", command=self._apply_theme).grid(row=0, column=2, sticky="w", padx=(10, 0))

        ttk.Checkbutton(appearance, text="Use recursive folder scanning by default", variable=self.recursive_var).grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(14, 0)
        )
        ttk.Checkbutton(appearance, text="Enable first-launch splash", variable=self.splash_enabled_var).grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(8, 0)
        )
        ttk.Checkbutton(appearance, text="Enable login reminder popup", variable=self.login_popup_enabled_var).grid(
            row=3, column=0, columnspan=3, sticky="w", pady=(8, 0)
        )
        ttk.Checkbutton(appearance, text="Auto-open output folder after successful runs", variable=self.auto_open_output_var).grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(8, 0)
        )
        ttk.Checkbutton(appearance, text="Restore last session files and tool state on startup", variable=self.restore_session_var).grid(
            row=5, column=0, columnspan=3, sticky="w", pady=(8, 0)
        )
        ttk.Checkbutton(appearance, text="Clean temporary session files on exit", variable=self.cleanup_temp_var).grid(
            row=6, column=0, columnspan=3, sticky="w", pady=(8, 0)
        )
        ttk.Checkbutton(appearance, text="Enable update checker placeholder reminders", variable=self.update_checker_enabled_var).grid(
            row=7, column=0, columnspan=3, sticky="w", pady=(8, 0)
        )
        ttk.Label(
            appearance,
            textvariable=self.login_popup_state_var,
            style="CardBody.TLabel",
            wraplength=460,
            justify="left",
        ).grid(row=8, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Label(
            appearance,
            text="Theme, output folder, engine choice, splash options, reminder state, and session behavior are saved automatically when you close the app.",
            style="CardBody.TLabel",
            wraplength=460,
            justify="left",
        ).grid(row=9, column=0, columnspan=3, sticky="w", pady=(12, 0))

        ttk.Label(appearance, text="Performance mode:", style="Surface.TLabel").grid(row=10, column=0, sticky="w", pady=(14, 0))
        ttk.Combobox(
            appearance,
            textvariable=self.performance_mode_var,
            values=["eco", "balanced", "quality"],
            state="readonly",
            width=14,
        ).grid(row=10, column=1, sticky="w", padx=(8, 0), pady=(14, 0))
        ttk.Label(
            appearance,
            text="Eco caps heavy image rendering for older machines. Balanced is the default. Quality keeps your requested scale untouched.",
            style="CardBody.TLabel",
            wraplength=460,
            justify="left",
        ).grid(row=11, column=0, columnspan=3, sticky="w", pady=(8, 0))

        engine_frame = ttk.LabelFrame(content, text="Conversion engine")
        engine_frame.grid(row=0, column=1, sticky="nsew", pady=(0, 10))
        engine_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(engine_frame, text="Engine:", style="Surface.TLabel").grid(row=0, column=0, sticky="w")
        self.engine_combo = ttk.Combobox(engine_frame, textvariable=self.engine_mode_var, values=ENGINE_ORDER, state="readonly", width=18)
        self.engine_combo.grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.engine_combo.bind("<<ComboboxSelected>>", self._on_engine_changed)
        ttk.Button(engine_frame, text="Refresh status", command=self._refresh_dependency_status).grid(row=0, column=2, sticky="e", padx=(10, 0))

        ttk.Label(engine_frame, textvariable=self.engine_help_var, style="CardBody.TLabel", wraplength=460, justify="left").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(10, 0)
        )

        ttk.Label(engine_frame, text="LibreOffice soffice path:", style="Surface.TLabel").grid(row=2, column=0, sticky="w", pady=(12, 0))
        self.soffice_entry = ttk.Entry(engine_frame, textvariable=self.soffice_path_var)
        self.soffice_entry.grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=(12, 0))
        soffice_actions = ttk.Frame(engine_frame)
        soffice_actions.grid(row=2, column=2, sticky="e", pady=(12, 0))
        ttk.Button(soffice_actions, text="Browse", command=self._browse_soffice_path).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(soffice_actions, text="Test", command=self._test_soffice_path).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(soffice_actions, text="Clear", command=self._clear_soffice_path).grid(row=0, column=2)

        ttk.Label(engine_frame, text="Tesseract path:", style="Surface.TLabel").grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(engine_frame, textvariable=self.tesseract_path_var).grid(row=3, column=1, sticky="ew", padx=(8, 8), pady=(10, 0))
        tesseract_actions = ttk.Frame(engine_frame)
        tesseract_actions.grid(row=3, column=2, sticky="e", pady=(10, 0))
        ttk.Button(tesseract_actions, text="Browse", command=self._browse_tesseract_path).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(tesseract_actions, text="Test", command=self._test_tesseract_path).grid(row=0, column=1)

        ttk.Label(engine_frame, text="Active route:", style="Surface.TLabel").grid(row=4, column=0, sticky="w", pady=(12, 0))
        ttk.Label(engine_frame, textvariable=self.active_engine_var, style="StatusGood.TLabel", wraplength=460, justify="left").grid(
            row=4, column=1, columnspan=2, sticky="w", pady=(12, 0)
        )
        ttk.Label(engine_frame, textvariable=self.dependency_var, style="CardBody.TLabel", wraplength=460, justify="left").grid(
            row=5, column=0, columnspan=3, sticky="w", pady=(12, 0)
        )
        ttk.Label(
            engine_frame,
            text="Routing preview:",
            style="Surface.TLabel",
        ).grid(row=6, column=0, sticky="w", pady=(12, 0))
        ttk.Label(engine_frame, textvariable=self.route_preview_var, style="CardBody.TLabel", wraplength=460, justify="left").grid(
            row=7, column=0, columnspan=3, sticky="w", pady=(6, 0)
        )

        files_frame = ttk.LabelFrame(content, text="Files and folders")
        files_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        files_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(files_frame, text="State file:", style="Surface.TLabel").grid(row=0, column=0, sticky="w")
        self.state_path_entry = ttk.Entry(files_frame)
        self.state_path_entry.grid(row=0, column=1, sticky="ew", padx=(8, 8))
        self.state_path_entry.insert(0, str(APP_STATE_PATH))
        self.state_path_entry.state(["readonly"])
        ttk.Button(files_frame, text="Open folder", command=lambda: open_path(APP_STATE_PATH.parent)).grid(row=0, column=2, sticky="e")

        ttk.Label(files_frame, text="Markdown notes file:", style="Surface.TLabel").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.notes_path_entry = ttk.Entry(files_frame)
        self.notes_path_entry.grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=(10, 0))
        self.notes_path_entry.insert(0, str(self.notes_path))
        self.notes_path_entry.state(["readonly"])
        ttk.Button(files_frame, text="Open file", command=lambda: open_path(self.notes_path)).grid(row=1, column=2, sticky="e", pady=(10, 0))

        ttk.Label(files_frame, text="About profile file:", style="Surface.TLabel").grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.profile_path_entry = ttk.Entry(files_frame)
        self.profile_path_entry.grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=(10, 0))
        self.profile_path_entry.insert(0, str(self.about_profile_path))
        self.profile_path_entry.state(["readonly"])
        ttk.Button(files_frame, text="Edit JSON", command=lambda: open_path(self.about_profile_path)).grid(row=2, column=2, sticky="e", pady=(10, 0))

        ttk.Label(files_frame, text="Installer build notes:", style="Surface.TLabel").grid(row=3, column=0, sticky="w", pady=(10, 0))
        self.build_notes_entry = ttk.Entry(files_frame)
        self.build_notes_entry.grid(row=3, column=1, sticky="ew", padx=(8, 8), pady=(10, 0))
        self.build_notes_entry.insert(0, str(self.build_notes_path))
        self.build_notes_entry.state(["readonly"])
        ttk.Button(files_frame, text="Open file", command=lambda: open_path(self.build_notes_path)).grid(row=3, column=2, sticky="e", pady=(10, 0))

        ttk.Label(files_frame, text="Splash GIF asset:", style="Surface.TLabel").grid(row=4, column=0, sticky="w", pady=(10, 0))
        self.splash_path_entry = ttk.Entry(files_frame, textvariable=self.splash_gif_path_var)
        self.splash_path_entry.grid(row=4, column=1, sticky="ew", padx=(8, 8), pady=(10, 0))
        splash_actions = ttk.Frame(files_frame)
        splash_actions.grid(row=4, column=2, sticky="e", pady=(10, 0))
        ttk.Button(splash_actions, text="Browse", command=self._browse_splash_gif_path).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(splash_actions, text="Preview", command=self._preview_splash).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(splash_actions, text="Reset", command=self._clear_splash_gif_path).grid(row=0, column=2)

        ttk.Label(files_frame, text="Static installer About file:", style="Surface.TLabel").grid(row=5, column=0, sticky="w", pady=(10, 0))
        self.static_about_entry = ttk.Entry(files_frame)
        self.static_about_entry.grid(row=5, column=1, sticky="ew", padx=(8, 8), pady=(10, 0))
        self.static_about_entry.insert(0, str(self.static_about_profile_path))
        self.static_about_entry.state(["readonly"])
        ttk.Button(files_frame, text="Open file", command=lambda: open_path(self.static_about_profile_path)).grid(row=5, column=2, sticky="e", pady=(10, 0))

        ttk.Label(files_frame, text="Install date:", style="Surface.TLabel").grid(row=6, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(files_frame, textvariable=self.install_date_var, state="readonly").grid(row=6, column=1, sticky="ew", padx=(8, 8), pady=(10, 0))
        ttk.Button(files_frame, text="Open state folder", command=lambda: open_path(APP_STATE_PATH.parent)).grid(row=6, column=2, sticky="e", pady=(10, 0))

        ttk.Label(files_frame, text="Last update check:", style="Surface.TLabel").grid(row=7, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(files_frame, textvariable=self.last_update_check_var, state="readonly").grid(row=7, column=1, sticky="ew", padx=(8, 8), pady=(10, 0))
        ttk.Button(files_frame, text="Check now", command=self._check_for_updates_placeholder).grid(row=7, column=2, sticky="e", pady=(10, 0))

        ttk.Label(files_frame, text="Link cache folder:", style="Surface.TLabel").grid(row=8, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(files_frame, textvariable=self.link_cache_dir_var).grid(row=8, column=1, sticky="ew", padx=(8, 8), pady=(10, 0))
        link_cache_actions = ttk.Frame(files_frame)
        link_cache_actions.grid(row=8, column=2, sticky="e", pady=(10, 0))
        ttk.Button(link_cache_actions, text="Open", command=self._open_link_cache_dir).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(link_cache_actions, text="Clear", command=self._clear_link_cache).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(link_cache_actions, text="Prune", command=self._prune_link_cache).grid(row=0, column=2)

        ttk.Label(files_frame, textvariable=self.link_cache_summary_var, style="CardBody.TLabel", wraplength=460, justify="left").grid(
            row=9, column=0, columnspan=3, sticky="w", pady=(8, 0)
        )

        ttk.Label(files_frame, text="Cache keep days:", style="Surface.TLabel").grid(row=10, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(files_frame, textvariable=self.link_cache_max_age_var, width=10).grid(row=10, column=1, sticky="w", padx=(8, 8), pady=(10, 0))
        ttk.Checkbutton(files_frame, text="Keep downloaded link files in cache", variable=self.link_keep_downloads_var).grid(
            row=10, column=2, sticky="e", pady=(10, 0)
        )

        ttk.Label(files_frame, text="Cache size cap (MB):", style="Surface.TLabel").grid(row=11, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(files_frame, textvariable=self.link_cache_max_size_var, width=10).grid(row=11, column=1, sticky="w", padx=(8, 8), pady=(10, 0))
        ttk.Button(files_frame, text="Refresh cache stats", command=self._refresh_link_cache_summary).grid(row=11, column=2, sticky="e", pady=(10, 0))

        ttk.Label(files_frame, text="Link timeout (seconds):", style="Surface.TLabel").grid(row=12, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(files_frame, textvariable=self.link_timeout_var).grid(row=12, column=1, sticky="w", padx=(8, 8), pady=(10, 0))

        ttk.Label(
            files_frame,
            text="Edit footer_notes.md or about_profile.json later, export settings snapshots when you want backups, and open installer/BUILDING.md when you want to prepare a packaged build.",
            style="CardBody.TLabel",
            wraplength=460,
            justify="left",
        ).grid(row=13, column=0, columnspan=3, sticky="w", pady=(14, 0))

        actions_frame = ttk.LabelFrame(content, text="Quick actions")
        actions_frame.grid(row=1, column=1, sticky="nsew")
        actions_frame.grid_columnconfigure(0, weight=1)
        ttk.Button(actions_frame, text="Open OCR page", command=lambda: self._show_page("ocr")).grid(row=0, column=0, sticky="ew")
        ttk.Button(actions_frame, text="Quick actions palette", command=self._open_command_palette).grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Preview splash", command=self._preview_splash).grid(row=2, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Reset login reminder", command=self._reset_login_popup_state).grid(row=3, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Test Tesseract", command=self._test_tesseract_path).grid(row=4, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Open organizer", command=lambda: self._show_page("organizer")).grid(row=5, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Open footer notes", command=self._open_notes_window).grid(row=6, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Open About page", command=self._show_about).grid(row=7, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Edit About profile", command=self._open_about_editor_window).grid(row=8, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Open SMTP delivery", command=self._open_smtp_window).grid(row=9, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Open link cache", command=self._open_link_cache_dir).grid(row=10, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Clear link cache", command=self._clear_link_cache).grid(row=11, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Send latest outputs", command=self._send_last_outputs_via_smtp).grid(row=12, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Save EML draft", command=self._save_eml_draft_for_last_outputs).grid(row=13, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Open Build Center", command=self._open_build_center_window).grid(row=14, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Open build notes", command=lambda: open_path(self.build_notes_path)).grid(row=15, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Export settings snapshot", command=self._export_state_snapshot_action).grid(row=16, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Import settings snapshot", command=self._import_state_snapshot_action).grid(row=17, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Export diagnostics report", command=self._export_diagnostics_report_action).grid(row=18, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Export app logs", command=self._export_current_logs_action).grid(row=19, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Check for updates", command=self._check_for_updates_placeholder).grid(row=20, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Reload About profile", command=self._refresh_about_profile).grid(row=20, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Refresh history view", command=self._refresh_history_views).grid(row=21, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions_frame, text="Clear local history", command=self._clear_history).grid(row=20, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(actions_frame, textvariable=self.smtp_status_var, style="CardBody.TLabel", wraplength=320, justify="left").grid(row=21, column=0, sticky="w", pady=(12, 0))


    def _build_about_page(self) -> None:
        page = ttk.Frame(self.content)
        page.grid(row=0, column=0, sticky="nsew")
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(1, weight=1)
        self.pages["about"] = page

        header = ttk.Frame(page, style="Card.TFrame", padding=22)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        ttk.Label(header, text="About Gokul Omni Convert Lite", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text=(
                "This page reads your editable profile JSON, keeps your in-app About section customizable, "
                "and also mirrors a static installer-safe snapshot for later packaging."
            ),
            style="CardBody.TLabel",
            wraplength=980,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        body = ttk.Frame(page)
        body.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=2)
        body.grid_rowconfigure(0, weight=1)

        profile_card = ttk.Frame(body, style="Card.TFrame", padding=20)
        profile_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        profile_card.grid_columnconfigure(0, weight=1)
        self.about_image_label = ttk.Label(profile_card, text="Loading profile image...")
        self.about_image_label.grid(row=0, column=0, sticky="n", pady=(0, 12))
        self.about_image_hint_label = ttk.Label(profile_card, style="CardBody.TLabel", wraplength=280, justify="center")
        self.about_image_hint_label.grid(row=1, column=0, sticky="ew")
        profile_actions = ttk.Frame(profile_card)
        profile_actions.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        profile_actions.grid_columnconfigure(0, weight=1)
        ttk.Button(profile_actions, text="Reload profile", command=self._refresh_about_profile).grid(row=0, column=0, sticky="ew")
        ttk.Button(profile_actions, text="Edit profile in app", command=self._open_about_editor_window).grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(profile_actions, text="Edit profile JSON", command=lambda: open_path(self.about_profile_path)).grid(row=2, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(profile_actions, text="Open image file", command=self._open_about_image_file).grid(row=3, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(profile_actions, text="Open static installer About", command=lambda: open_path(self.static_about_profile_path)).grid(row=4, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(profile_actions, text="Open build notes", command=lambda: open_path(self.build_notes_path)).grid(row=5, column=0, sticky="ew", pady=(8, 0))

        info_card = ttk.Frame(body, style="Card.TFrame", padding=22)
        info_card.grid(row=0, column=1, sticky="nsew")
        info_card.grid_columnconfigure(0, weight=1)

        self.about_name_label = ttk.Label(info_card, style="Title.TLabel")
        self.about_name_label.grid(row=0, column=0, sticky="w")
        self.about_title_label = ttk.Label(info_card, style="Subtitle.TLabel")
        self.about_title_label.grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.about_company_label = ttk.Label(info_card, style="CardBody.TLabel", wraplength=680, justify="left")
        self.about_company_label.grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.about_meta_label = ttk.Label(info_card, style="CardBody.TLabel", wraplength=680, justify="left")
        self.about_meta_label.grid(row=3, column=0, sticky="w", pady=(8, 0))
        self.about_bio_label = ttk.Label(info_card, style="CardBody.TLabel", wraplength=680, justify="left")
        self.about_bio_label.grid(row=4, column=0, sticky="w", pady=(12, 0))

        ttk.Separator(info_card, orient="horizontal").grid(row=5, column=0, sticky="ew", pady=14)
        ttk.Label(info_card, text="Actions", style="CardTitle.TLabel").grid(row=6, column=0, sticky="w")
        self.about_action_frame = ttk.Frame(info_card)
        self.about_action_frame.grid(row=7, column=0, sticky="w", pady=(10, 0))

        ttk.Separator(info_card, orient="horizontal").grid(row=8, column=0, sticky="ew", pady=14)
        ttk.Label(info_card, text="Social links", style="CardTitle.TLabel").grid(row=9, column=0, sticky="w")
        self.about_links_frame = ttk.Frame(info_card)
        self.about_links_frame.grid(row=10, column=0, sticky="w", pady=(10, 0))

        ttk.Separator(info_card, orient="horizontal").grid(row=11, column=0, sticky="ew", pady=14)
        ttk.Label(info_card, text=f"Version: {APP_VERSION}", style="CardBody.TLabel").grid(row=12, column=0, sticky="w")
        ttk.Label(
            info_card,
            text="Installer builds should read from the static snapshot in installer/about_static.json. The app itself continues to read the editable local profile assets.",
            style="CardBody.TLabel",
            wraplength=680,
            justify="left",
        ).grid(row=13, column=0, sticky="w", pady=(8, 0))
        ttk.Label(
            info_card,
            text=f"State file: {APP_STATE_PATH}",
            style="CardBody.TLabel",
            wraplength=680,
            justify="left",
        ).grid(row=14, column=0, sticky="w", pady=(8, 0))
    def _open_about_image_file(self) -> None:
        profile = load_about_profile(self.about_profile_path)
        image_path = resolve_profile_image(profile, self.about_profile_path.parent)
        if image_path.exists():
            open_path(image_path)
        else:
            open_path(self.about_profile_path.parent)


    def _refresh_about_profile(self) -> None:
        self.about_profile = load_about_profile(self.about_profile_path)
        profile = self.about_profile
        self.engine_help_var.set(ENGINE_HELP.get(self.engine_mode_var.get().strip().lower() or ENGINE_AUTO, ""))
        self.active_engine_var.set((self.engine_mode_var.get().strip() or ENGINE_PURE_PYTHON).replace("_", " ").title())
        self._sync_static_about_profile()

        if not self.smtp_sender_var.get().strip() and str(profile.get("email", "")).strip():
            self.smtp_sender_var.set(str(profile.get("email", "")).strip())

        if hasattr(self, "about_name_label"):
            self.about_name_label.configure(text=str(profile.get("name", APP_NAME)).strip() or APP_NAME)
            subtitle = str(profile.get("title", "")).strip()
            extra = str(profile.get("subtitle", "")).strip()
            subtitle_text = subtitle if not extra else f"{subtitle} • {extra}" if subtitle else extra
            self.about_title_label.configure(text=subtitle_text)

            company = str(profile.get("company", "")).strip()
            project = str(profile.get("project", "")).strip()
            company_parts = [item for item in [company, project] if item]
            self.about_company_label.configure(text=" • ".join(company_parts))

            meta_parts = []
            if str(profile.get("email", "")).strip():
                meta_parts.append(f"Email: {str(profile.get('email', '')).strip()}")
            if str(profile.get("handle", "")).strip():
                meta_parts.append(f"Handle: {str(profile.get('handle', '')).strip()}")
            self.about_meta_label.configure(text=" | ".join(meta_parts))
            self.about_bio_label.configure(text=str(profile.get("bio", "")).strip())

            for child in self.about_action_frame.winfo_children():
                child.destroy()
            action_buttons: list[tuple[str, str]] = []
            email = str(profile.get("email", "")).strip()
            feedback_url = str(profile.get("feedback_url", "")).strip()
            contribute_url = str(profile.get("contribute_url", "")).strip()
            if email:
                action_buttons.append(("Email", f"mailto:{email}"))
            if feedback_url:
                action_buttons.append(("Feedback", feedback_url))
            if contribute_url:
                action_buttons.append(("Contribute", contribute_url))
            if not action_buttons:
                ttk.Label(self.about_action_frame, text="No primary actions configured yet.", style="CardBody.TLabel").grid(row=0, column=0, sticky="w")
            else:
                for index, (label, url) in enumerate(action_buttons):
                    ttk.Button(self.about_action_frame, text=label, command=lambda target=url: open_url(target)).grid(
                        row=0, column=index, padx=(0 if index == 0 else 8, 0), sticky="w"
                    )

            for child in self.about_links_frame.winfo_children():
                child.destroy()
            links = profile.get("links", []) if isinstance(profile.get("links"), list) else []
            active_links = [item for item in links if isinstance(item, dict) and str(item.get("url", "")).strip()]
            if not active_links:
                ttk.Label(
                    self.about_links_frame,
                    text="No social links configured yet. Use the in-app editor or edit about_profile.json to add them.",
                    style="CardBody.TLabel",
                ).grid(row=0, column=0, sticky="w")
            else:
                for index, item in enumerate(active_links):
                    label = str(item.get("label", "Link")).strip() or "Link"
                    url = str(item.get("url", "")).strip()
                    ttk.Button(self.about_links_frame, text=label, command=lambda target=url: open_url(target)).grid(
                        row=0, column=index, padx=(0 if index == 0 else 8, 0), sticky="w"
                    )

            image_path = resolve_profile_image(profile, self.about_profile_path.parent)
            self.about_image_hint_label.configure(text=f"Profile image: {image_path.name}")
            if image_path.exists():
                try:
                    image = Image.open(image_path)
                    image.thumbnail((280, 280))
                    self.about_photo = ImageTk.PhotoImage(image)
                    self.about_image_label.configure(image=self.about_photo, text="")
                except Exception:
                    self.about_photo = None
                    self.about_image_label.configure(image="", text="Could not load the profile image.")
            else:
                self.about_photo = None
                self.about_image_label.configure(image="", text="Profile image not found.")

        if self.about_editor_window and self.about_editor_window.winfo_exists():
            self.about_editor_window.load_profile()
        if self.build_center_window and self.build_center_window.winfo_exists():
            self.build_center_window.refresh_summary()
    def _browse_soffice_path(self) -> None:
        initial = self.soffice_path_var.get().strip() or str(Path.home())
        file_path = filedialog.askopenfilename(
            title="Select soffice executable",
            initialdir=str(Path(initial).expanduser().parent if Path(initial).expanduser().exists() else Path.home()),
            filetypes=[("Executable", "soffice*"), ("All files", "*.*")],
        )
        if file_path:
            self.soffice_path_var.set(file_path)
            self._refresh_dependency_status()

    def _test_soffice_path(self) -> None:
        self._refresh_dependency_status()
        status = dependency_status(self.soffice_path_var.get().strip())
        if status.get("LibreOffice"):
            messagebox.showinfo("Soffice", "LibreOffice / soffice was found successfully.")
        else:
            messagebox.showwarning("Soffice", "LibreOffice / soffice was not found. Check the configured path or install LibreOffice.")

    def _browse_tesseract_path(self) -> None:
        initial = self.tesseract_path_var.get().strip() or str(Path.home())
        file_path = filedialog.askopenfilename(
            title="Select Tesseract executable",
            initialdir=str(Path(initial).expanduser().parent if Path(initial).expanduser().exists() else Path.home()),
            filetypes=[("Executable", "tesseract*"), ("All files", "*.*")],
        )
        if file_path:
            self.tesseract_path_var.set(file_path)
            self._refresh_dependency_status()

    def _test_tesseract_path(self) -> None:
        status = detect_tesseract_status(self.tesseract_path_var.get().strip())
        if bool(status.get("available")):
            summary = f"Tesseract found via {status.get('source')}: {status.get('path')}"
            self.ocr_status_var.set(summary)
            self.status_var.set("Tesseract is available.")
            messagebox.showinfo("Tesseract", summary)
        else:
            summary = "Tesseract was not found. Install it or configure its executable path in Settings."
            self.ocr_status_var.set(summary)
            self.status_var.set("Tesseract is missing.")
            messagebox.showwarning("Tesseract", summary)
        self._refresh_dependency_status()

    def _append_ocr_log(self, message: str) -> None:
        self.ocr_log_text.configure(state="normal")
        self.ocr_log_text.insert(tk.END, message.rstrip() + "\n")
        self.ocr_log_text.see(tk.END)
        self.ocr_log_text.configure(state="disabled")

    def _clear_ocr_log(self) -> None:
        self.ocr_log_text.configure(state="normal")
        self.ocr_log_text.delete("1.0", tk.END)
        self.ocr_log_text.configure(state="disabled")
        self.ocr_status_var.set("OCR log cleared.")

    def _browse_ocr_input(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select image or PDF for OCR",
            filetypes=[("Image or PDF", "*.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp *.pdf"), ("All files", "*.*")],
        )
        if file_path:
            self.ocr_input_var.set(file_path)

    def _browse_ocr_output(self) -> None:
        directory = filedialog.askdirectory(title="Select OCR output folder")
        if directory:
            self.ocr_output_var.set(directory)

    def _ocr_output_path(self, suffix: str) -> Path:
        source = Path(self.ocr_input_var.get().strip() or "ocr_input")
        output_dir = Path(self.ocr_output_var.get().strip() or (Path(self.output_dir_var.get().strip() or Path.cwd() / "converted_output") / "ocr"))
        output_dir.mkdir(parents=True, exist_ok=True)
        return output_dir / f"{source.stem}_{suffix}"

    def _ocr_config_from_vars(self) -> OcrConfig:
        try:
            dpi = int(str(self.ocr_dpi_var.get()).strip() or "220")
            psm = int(str(self.ocr_psm_var.get()).strip() or "6")
        except Exception as exc:
            raise ValueError("OCR DPI and PSM must be whole numbers.") from exc
        return OcrConfig(
            language=self.ocr_language_var.get().strip() or "eng",
            dpi=max(dpi, 72),
            psm=max(psm, 0),
            tesseract_path=self.tesseract_path_var.get().strip(),
        )

    def _set_ocr_buttons_enabled(self, enabled: bool) -> None:
        widgets = [
            getattr(self, "start_button", None),
            getattr(self, "pdf_tool_start_button", None),
            getattr(self, "ocr_start_image_button", None),
            getattr(self, "ocr_start_pdf_button", None),
            getattr(self, "ocr_start_text_button", None),
        ]
        state = ["!disabled"] if enabled else ["disabled"]
        for widget in widgets:
            if widget is None:
                continue
            try:
                widget.state(state)
            except Exception:
                pass

    def _start_ocr_job(self, label: str, runner, *, expected_pdf: bool | None = None) -> None:
        if self.running:
            return
        source = Path(self.ocr_input_var.get().strip()).expanduser()
        if not source.exists():
            messagebox.showerror("OCR", "Please choose an OCR input image or PDF first.")
            self._show_page("ocr")
            return
        if expected_pdf is True and source.suffix.lower() != ".pdf":
            messagebox.showwarning("OCR", "Select a PDF file for this OCR action.")
            self._show_page("ocr")
            return
        if expected_pdf is False and source.suffix.lower() == ".pdf":
            messagebox.showwarning("OCR", "Select an image file for this OCR action, or use PDF -> Searchable PDF.")
            self._show_page("ocr")
            return

        try:
            config = self._ocr_config_from_vars()
        except Exception as exc:
            messagebox.showerror("OCR settings", str(exc))
            self._show_page("ocr")
            return

        output_dir = Path(self.ocr_output_var.get().strip() or (Path(self.output_dir_var.get().strip() or Path.cwd() / "converted_output") / "ocr"))
        output_dir.mkdir(parents=True, exist_ok=True)

        record_base = {
            "job_type": "ocr",
            "mode": label,
            "file_count": 1,
            "output_count": 0,
            "output_dir": str(output_dir),
            "input_files": [str(source)],
            "inputs_preview": [str(source)],
            "outputs_preview": [],
            "engine_mode": "ocr",
            "merge_to_one_pdf": False,
            "merged_output_name": "",
            "image_format": "",
            "image_scale": "",
            "soffice_configured": False,
            "ocr_language": config.language,
            "ocr_dpi": config.dpi,
            "ocr_psm": config.psm,
            "tesseract_configured": bool(config.tesseract_path),
        }

        self.running = True
        self.active_run_kind = "ocr"
        self.last_run_origin = "ocr"
        self._set_ocr_buttons_enabled(False)
        self.ocr_progress["value"] = 0
        self.status_var.set("Running OCR...")
        self.ocr_status_var.set(f"Running: {label}")
        self._append_ocr_log(f"=== {label} ===")
        self._append_ocr_log(f"Input: {source}")
        self._append_ocr_log(f"Output folder: {output_dir}")
        if config.tesseract_path:
            self._append_ocr_log(f"Tesseract path: {config.tesseract_path}")

        def worker() -> None:
            try:
                progress_cb = lambda current, total: self.worker_queue.put(("ocr_progress", (current, total)))
                log_cb = lambda message: self.worker_queue.put(("ocr_log", message))
                output_path = runner(source, config, progress_cb, log_cb)
                record = dict(record_base)
                record["output_count"] = 1
                record["outputs_preview"] = [str(output_path)]
                self.worker_queue.put(("ocr_done", (record, [str(output_path)])))
            except Exception as exc:
                self.worker_queue.put(("ocr_error", (dict(record_base), str(exc))))

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self._poll_worker_queue)
        self._show_page("ocr")

    def _start_ocr_image_pdf(self) -> None:
        self._start_ocr_job(
            "OCR: Image -> Searchable PDF",
            lambda source, config, progress_cb, log_cb: image_to_searchable_pdf(
                source,
                self._ocr_output_path("searchable.pdf"),
                config=config,
                progress=progress_cb,
                log=log_cb,
            ),
            expected_pdf=False,
        )

    def _start_ocr_pdf_pdf(self) -> None:
        self._start_ocr_job(
            "OCR: PDF -> Searchable PDF",
            lambda source, config, progress_cb, log_cb: pdf_to_searchable_pdf(
                source,
                self._ocr_output_path("ocr.pdf"),
                config=config,
                progress=progress_cb,
                log=log_cb,
            ),
            expected_pdf=True,
        )

    def _start_ocr_text(self) -> None:
        self._start_ocr_job(
            "OCR: Extract OCR Text",
            lambda source, config, progress_cb, log_cb: extract_text_with_ocr(
                source,
                self._ocr_output_path("ocr.txt"),
                config=config,
                progress=progress_cb,
                log=log_cb,
            ),
            expected_pdf=None,
        )

    def _on_engine_changed(self, _event=None) -> None:
        self.engine_help_var.set(ENGINE_HELP.get(self.engine_mode_var.get().strip().lower() or ENGINE_AUTO, ""))
        self.active_engine_var.set((self.engine_mode_var.get().strip() or ENGINE_PURE_PYTHON).replace("_", " ").title())
        self._refresh_dependency_status()

    def _create_metric_card(self, parent: ttk.Frame, column: int, title: str, variable: tk.StringVar) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=18)
        card.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 10, 0))
        card.grid_columnconfigure(0, weight=1)
        ttk.Label(card, text=title, style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(card, textvariable=variable, style="CardBody.TLabel", wraplength=280, justify="left").grid(row=1, column=0, sticky="w", pady=(8, 0))

    def _build_supported_text(self, mode: str) -> str:
        extensions = ", ".join(sorted(supported_extensions_for_mode(mode)))
        return f"Supported file extensions: {extensions}"

    def _refresh_route_preview(self) -> None:
        preview = build_conversion_route_preview(
            self.mode_var.get(),
            self.selected_files,
            engine_mode=self.engine_mode_var.get().strip().lower() or ENGINE_AUTO,
            soffice_path=self.soffice_path_var.get().strip(),
        )
        self.route_preview_var.set(preview)

    def _show_page(self, name: str) -> None:
        if name not in self.pages:
            return
        self.pages[name].tkraise()
        self.current_page = name
        for page_name, button in self.nav_buttons.items():
            button.configure(style="NavActive.TButton" if page_name == name else "Nav.TButton")
        if name == "history":
            self._refresh_history_views()
        if name == "home":
            self._refresh_home_summary()
            self._refresh_favorite_preset_widgets()
        if name == "settings":
            self._refresh_link_cache_summary()
        if name == "organizer" and self.organizer_panel is not None and self.organizer_panel.source_pdf is None:
            remembered = str(self.state_store.get("organizer_last_pdf", "")).strip()
            if remembered and Path(remembered).exists():
                self.organizer_panel.load_pdf(Path(remembered))

    def _on_mode_changed(self, _event=None) -> None:
        self.output_name_var.set(default_merged_name(self.mode_var.get()))
        self._update_mode_controls()

    def _effective_image_scale(self, value: float) -> float:
        mode = self.performance_mode_var.get().strip().lower()
        if mode == "eco":
            return max(1.0, min(float(value), 1.5))
        if mode == "balanced":
            return max(1.0, min(float(value), 2.0))
        return max(1.0, float(value))


    def _update_mode_controls(self) -> None:
        mode = self.mode_var.get()
        self.mode_help_var.set(MODE_HELP.get(mode, ""))
        self.supported_var.set(self._build_supported_text(mode))
        self.home_mode_var.set(mode)
        self._refresh_route_preview()

        pdf_output_mode = outputs_pdf(mode) and mode != MODE_MERGE_PDFS
        merge_name_enabled = pdf_output_mode or mode == MODE_MERGE_PDFS
        image_mode = mode in {MODE_PDF_TO_IMAGES, MODE_PDF_TO_PPTX, MODE_PRESENTATIONS_TO_IMAGES}

        if pdf_output_mode:
            self.merge_check.state(["!disabled"])
        else:
            self.merge_var.set(False)
            self.merge_check.state(["disabled"])

        if merge_name_enabled:
            self.output_name_entry.state(["!disabled"])
        else:
            self.output_name_entry.state(["disabled"])

        if image_mode:
            self.image_format_combo.state(["!disabled", "readonly"])
            self.image_scale_entry.state(["!disabled"])
        else:
            self.image_format_combo.state(["disabled"])
            self.image_scale_entry.state(["disabled"])

        self._refresh_home_summary()

    def _refresh_dependency_status(self) -> None:
        engine = self.engine_mode_var.get().strip().lower() or ENGINE_AUTO
        self.engine_help_var.set(ENGINE_HELP.get(engine, ""))
        self.active_engine_var.set(engine.replace("_", " ").title())

        configured_path = self.soffice_path_var.get().strip()
        status = dependency_status(configured_path)
        if configured_path:
            soffice_state = "configured path" if status.get("LibreOffice") else "configured path missing"
        else:
            soffice_state = "PATH" if status.get("LibreOffice") else "not configured"

        tesseract = detect_tesseract_status(self.tesseract_path_var.get().strip())
        tesseract_state = str(tesseract.get("source", "missing"))
        if tesseract_state == "configured":
            tesseract_state = "configured path" if tesseract.get("available") else "configured path missing"

        parts: list[str] = [
            "Built-in pure Python: Ready",
            f"LibreOffice: {'Found' if status.get('LibreOffice') else 'Missing'} ({soffice_state})",
            f"Tesseract: {'Found' if tesseract.get('available') else 'Missing'} ({tesseract_state})",
            f"Pandoc: {'Found' if status.get('Pandoc') else 'Missing'}",
            f"pdftotext: {'Found' if status.get('pdftotext') else 'Missing'}",
            f"python-pptx: {'Found' if status.get('python-pptx') else 'Missing'}",
            f"xlrd: {'Found' if status.get('xlrd') else 'Missing'}",
        ]
        self.dependency_var.set("Dependency status -> " + " | ".join(parts))
        if bool(tesseract.get("available")):
            self.ocr_dependency_var.set(f"Tesseract ready via {tesseract.get('source')}: {tesseract.get('path')}")
            if "missing" in self.ocr_status_var.get().lower():
                self.ocr_status_var.set("OCR tools are ready. Add an image or PDF to begin.")
        else:
            self.ocr_dependency_var.set("Tesseract is missing. Install it or configure its path in Settings or the OCR page.")
        self._update_login_popup_state()
        self._refresh_route_preview()
        self._refresh_home_summary()

    def _refresh_home_summary(self) -> None:
        count = len(self.selected_files)
        self.home_selected_count_var.set(f"{count} file{'s' if count != 1 else ''} selected")
        self.home_output_var.set(self.output_dir_var.get().strip() or str(Path.cwd() / "converted_output"))
        engine = self.engine_mode_var.get().strip().lower() or ENGINE_AUTO
        engine_label = {
            ENGINE_AUTO: "auto engine",
            ENGINE_PURE_PYTHON: "pure Python engine",
            ENGINE_LIBREOFFICE: "LibreOffice engine",
        }.get(engine, engine)
        if self.running:
            self.home_hint_var.set("A job is currently running. Watch progress in the Convert, PDF Tools, or OCR screen.")
        elif count == 0:
            if hasattr(self, "link_input_text") and self._extract_urls_from_text():
                self.home_hint_var.set("URLs are ready in the online links area. Fetch them or press Start conversion to download and run them.")
            else:
                self.home_hint_var.set("No inputs are loaded yet. Add files, a folder, or online links from Home or Convert.")
        else:
            merge_state = "merged output" if self.merge_var.get() else "separate outputs"
            self.home_hint_var.set(f"Ready to run {self.mode_var.get()} using {engine_label} with {merge_state}.")

    def _add_files(self) -> None:
        mode = self.mode_var.get()
        patterns = filetype_patterns_for_mode(mode)
        file_paths = filedialog.askopenfilenames(
            title="Select input files",
            filetypes=[(f"Supported files for {mode}", patterns), ("All files", "*.*")],
        )
        if not file_paths:
            return
        self._append_files([Path(path) for path in file_paths])
        self._show_page("convert")

    def _add_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select input folder")
        if not folder:
            return

        files = collect_files_from_folder(
            Path(folder),
            supported_extensions_for_mode(self.mode_var.get()),
            recursive=self.recursive_var.get(),
        )
        if not files:
            messagebox.showinfo("No matching files", "No supported files were found in that folder for the current mode.")
            return
        self._append_files(files)
        self._show_page("convert")

    def _append_files(self, new_files: list[Path]) -> None:
        merged = {str(path.resolve()): path for path in self.selected_files}
        for path in new_files:
            merged[str(path.resolve())] = Path(path).resolve()
        self.selected_files = sorted(merged.values(), key=lambda p: str(p).lower())
        self._refresh_file_listbox()
        self.status_var.set(f"Selected {len(self.selected_files)} file(s).")
        self._refresh_route_preview()
        self._refresh_home_summary()

    def _refresh_file_listbox(self) -> None:
        self.file_listbox.delete(0, tk.END)
        for path in self.selected_files:
            self.file_listbox.insert(tk.END, str(path))


    def _extract_urls_from_text(self) -> list[str]:
        return extract_urls(self.link_input_text.get("1.0", tk.END))

    def _refresh_recent_links_summary(self) -> None:
        recent = self.state_store.recent_links()
        if recent:
            preview = "  |  ".join(recent[:3])
            if len(recent) > 3:
                preview += f"  |  ... and {len(recent) - 3} more"
            self.link_recent_summary_var.set(f"Recent links: {preview}")
        else:
            self.link_recent_summary_var.set("Recent links: none yet")
        downloaded = sum(1 for item in self.link_status_items.values() if str(item.get("status", "")) == "downloaded")
        self.link_fetch_count_var.set(f"{downloaded} downloaded")

    def _paste_urls_from_clipboard(self) -> None:
        try:
            clipboard = self.clipboard_get()
        except Exception:
            clipboard = ""
        if not clipboard.strip():
            self.status_var.set("Clipboard does not contain any text URLs.")
            return
        existing = self.link_input_text.get("1.0", tk.END).strip()
        if existing:
            self.link_input_text.insert(tk.END, ("\n" if not existing.endswith("\n") else "") + clipboard.strip() + "\n")
        else:
            self.link_input_text.insert("1.0", clipboard.strip() + "\n")
        self.status_var.set("Pasted URLs from the clipboard.")

    def _load_recent_links(self) -> None:
        recent = self.state_store.recent_links()
        if not recent:
            self.status_var.set("No recent links are stored yet.")
            return
        self.link_input_text.delete("1.0", tk.END)
        self.link_input_text.insert("1.0", "\n".join(recent[:20]) + "\n")
        self.status_var.set("Loaded recent links into the URL box.")

    def _clear_link_urls(self) -> None:
        self.link_input_text.delete("1.0", tk.END)
        self.link_status_items.clear()
        for item in self.link_status_tree.get_children():
            self.link_status_tree.delete(item)
        self.link_downloaded_files = []
        self.link_status_summary_var.set("Paste one or more direct file links or page URLs, then fetch them into the same conversion queue.")
        self._refresh_recent_links_summary()

    def _resolve_link_cache_dir(self) -> Path:
        return cache_root_from_setting(self.link_cache_dir_var.get().strip(), APP_STATE_PATH.parent)

    def _open_link_cache_dir(self) -> None:
        open_path(self._resolve_link_cache_dir())

    def _refresh_link_cache_summary(self) -> None:
        try:
            age_days = max(0, int(str(self.link_cache_max_age_var.get()).strip() or "0"))
        except Exception:
            age_days = 0
            self.link_cache_max_age_var.set("0")
        try:
            size_mb = max(32, int(str(self.link_cache_max_size_var.get()).strip() or "512"))
        except Exception:
            size_mb = 512
            self.link_cache_max_size_var.set("512")
        summary = summarize_directory(self._resolve_link_cache_dir())
        self.link_cache_summary_var.set(f"{summary} • policy: keep <= {size_mb} MB, prune > {age_days} day(s).")

    def _clear_link_cache(self) -> None:
        if self.running:
            return
        file_count, byte_count = clear_cache_dir(self._resolve_link_cache_dir())
        self._refresh_link_cache_summary()
        self.status_var.set(f"Cleared link cache ({file_count} files, {byte_count // 1024} KB).")

    def _prune_link_cache(self) -> None:
        if self.running:
            return
        try:
            age_days = max(0, int(str(self.link_cache_max_age_var.get()).strip() or "0"))
            size_mb = max(32, int(str(self.link_cache_max_size_var.get()).strip() or "512"))
        except Exception:
            messagebox.showerror("Link cache policy", "Cache age and size values must be whole numbers.")
            return
        result = prune_directory(
            self._resolve_link_cache_dir(),
            max_age_days=age_days,
            max_total_bytes=size_mb * 1024 * 1024,
        )
        self._refresh_link_cache_summary()
        self.status_var.set(
            f"Pruned cache: removed {result.get('removed_count', 0)} file(s), {format_bytes(int(result.get('removed_bytes', 0)))}."
        )

    def _pause_link_fetch(self) -> None:
        if self.active_run_kind != "links":
            self.status_var.set("No link download is running.")
            return
        self.link_pause_event.set()
        self._update_link_fetch_buttons()
        self.status_var.set("Link downloads paused.")

    def _resume_link_fetch(self) -> None:
        if self.active_run_kind != "links":
            self.status_var.set("No link download is running.")
            return
        self.link_pause_event.clear()
        self._update_link_fetch_buttons()
        self.status_var.set("Link downloads resumed.")

    def _update_link_fetch_buttons(self) -> None:
        pause_button = getattr(self, "link_pause_button", None)
        resume_button = getattr(self, "link_resume_button", None)
        if pause_button is not None:
            if self.active_run_kind == "links" and self.running and not self.link_pause_event.is_set():
                pause_button.state(["!disabled"])
            else:
                pause_button.state(["disabled"])
        if resume_button is not None:
            if self.active_run_kind == "links" and self.running and self.link_pause_event.is_set():
                resume_button.state(["!disabled"])
            else:
                resume_button.state(["disabled"])

    def _cancel_link_fetch(self) -> None:
        if self.active_run_kind != "links":
            self.status_var.set("No link download is running.")
            return
        self.link_cancel_event.set()
        self.link_pause_event.clear()
        self._update_link_fetch_buttons()
        self.status_var.set("Cancelling link downloads...")

    def _upsert_link_status(self, record: dict[str, object]) -> None:
        key = str(record.get("normalized_url") or record.get("url") or f"row_{len(self.link_status_items) + 1}")
        self.link_status_items[key] = record
        item_id = f"link_{abs(hash(key))}"
        values = (
            str(record.get("status", "")),
            str(record.get("url", ""))[:220],
            str(record.get("file_path", "") or record.get("filename", ""))[:180],
            str(record.get("detail", "") or record.get("error", ""))[:100],
        )
        if self.link_status_tree.exists(item_id):
            self.link_status_tree.item(item_id, values=values)
        else:
            self.link_status_tree.insert("", tk.END, iid=item_id, values=values)
        self._refresh_recent_links_summary()

    def _on_link_status_selected(self, _event=None) -> None:
        item = self.link_status_tree.focus()
        if not item:
            return
        values = self.link_status_tree.item(item, "values")
        if values:
            self.status_var.set(f"Link status -> {values[0]} | {values[1]}")

    def _start_link_fetch(self, auto_start: bool = False, retry_failed: bool = False) -> None:
        if self.running:
            return

        if retry_failed:
            urls = [
                str(item.get("url", "")).strip()
                for item in self.link_status_items.values()
                if str(item.get("status", "")).lower() in {"failed", "cancelled", "invalid"}
            ]
        else:
            urls = self._extract_urls_from_text()

        urls = [item for item in urls if item]
        if not urls:
            messagebox.showwarning("No URLs found", "Paste one or more http(s) URLs first.")
            self._show_page("convert")
            return

        try:
            timeout = max(int(str(self.link_timeout_var.get()).strip() or "25"), 5)
        except Exception:
            messagebox.showerror("Invalid timeout", "Link timeout must be a whole number of seconds.")
            return

        if not bool(self.link_keep_downloads_var.get()):
            self._clear_link_cache()

        self.running = True
        self.active_run_kind = "links"
        self.link_auto_start_pending = bool(auto_start)
        self.link_cancel_event.clear()
        self.link_pause_event.clear()
        self.progress["value"] = 0
        self.status_var.set("Fetching online links...")
        self._append_log("\n=== Link fetch started ===")
        self._append_log(f"Links queued: {len(urls)}")
        self._append_log(f"Cache folder: {self._resolve_link_cache_dir()}")
        self._append_log(f"Keep downloads: {bool(self.link_keep_downloads_var.get())}")
        self._append_log(f"Timeout: {timeout}s")
        self.start_button.state(["disabled"])
        self.pdf_tool_start_button.state(["disabled"])
        self._update_link_fetch_buttons()
        worker = threading.Thread(target=self._worker_link_fetch, args=(urls, timeout, bool(auto_start)), daemon=True)
        worker.start()
        self.after(100, self._poll_worker_queue)

    def _worker_link_fetch(self, urls: list[str], timeout: int, auto_start: bool) -> None:
        cache_dir = self._resolve_link_cache_dir()

        def status_update(result) -> None:
            self.worker_queue.put(("link_status", dict(vars(result))))

        def progress(current: int, total: int) -> None:
            self.worker_queue.put(("link_progress", (current, total)))

        try:
            results = download_many_urls(
                urls,
                cache_dir,
                timeout=timeout,
                cancel_requested=self.link_cancel_event.is_set,
                pause_requested=self.link_pause_event.is_set,
                status_callback=status_update,
                progress_callback=progress,
                keep_downloads=bool(self.link_keep_downloads_var.get()),
            )
            self.worker_queue.put(("link_done", (urls, results, auto_start)))
        except Exception:
            self.worker_queue.put(("link_error", traceback.format_exc()))

    def _remove_selected(self) -> None:
        selected_indexes = list(self.file_listbox.curselection())
        if not selected_indexes:
            return
        for index in reversed(selected_indexes):
            del self.selected_files[index]
        self._refresh_file_listbox()
        self.status_var.set(f"Selected {len(self.selected_files)} file(s).")
        self._refresh_route_preview()
        self._refresh_home_summary()

    def _clear_inputs(self) -> None:
        self.selected_files.clear()
        self._refresh_file_listbox()
        self.status_var.set("Ready.")
        self._refresh_route_preview()
        self._refresh_home_summary()

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")
        self.status_var.set("Activity log cleared.")

    def _browse_output_dir(self) -> None:
        folder = filedialog.askdirectory(title="Choose output folder")
        if folder:
            self.output_dir_var.set(folder)
            self._refresh_home_summary()

    def _open_output_folder(self) -> None:
        folder = Path(self.output_dir_var.get().strip()).expanduser()
        folder.mkdir(parents=True, exist_ok=True)
        open_path(folder)

    def _start_conversion(self) -> None:
        if self.running:
            return
        if not self.selected_files:
            url_candidates = self._extract_urls_from_text()
            if url_candidates:
                self._start_link_fetch(auto_start=True)
                return
            messagebox.showwarning("No files selected", "Please add some input files, paste URLs, or choose a folder first.")
            self._show_page("convert")
            return

        try:
            image_scale = float(self.image_scale_var.get().strip() or "2.0")
            image_scale = self._effective_image_scale(image_scale)
        except ValueError:
            messagebox.showerror("Invalid image scale", "Image scale must be a number such as 1.5 or 2.0.")
            self._show_page("convert")
            return

        output_dir_text = self.output_dir_var.get().strip()
        if not output_dir_text:
            output_dir = self.selected_files[0].parent / "converted_output"
            self.output_dir_var.set(str(output_dir))
        else:
            output_dir = Path(output_dir_text).expanduser()

        config = BatchConfig(
            mode=self.mode_var.get(),
            files=self.selected_files.copy(),
            output_dir=output_dir,
            merge_to_one_pdf=self.merge_var.get(),
            merged_output_name=self.output_name_var.get().strip() or default_merged_name(self.mode_var.get()),
            image_format=self.image_format_var.get().strip() or "png",
            image_scale=image_scale,
            engine_mode=self.engine_mode_var.get().strip().lower() or ENGINE_AUTO,
            soffice_path=self.soffice_path_var.get().strip(),
        )

        self.running = True
        self.active_run_kind = "convert"
        self._set_ocr_buttons_enabled(False)
        self.progress["value"] = 0
        self.status_var.set("Working...")
        self._append_log("\n=== New run started ===")
        self._append_log(f"Mode: {config.mode}")
        self._append_log(f"Engine: {config.engine_mode}")
        self._append_log(self.route_preview_var.get())
        if config.soffice_path:
            self._append_log(f"Configured soffice path: {config.soffice_path}")
        self._append_log(f"Output folder: {config.output_dir}")
        self._append_log(f"Files selected: {len(config.files)}")

        worker = threading.Thread(target=self._worker_run, args=(config,), daemon=True)
        worker.start()
        self.after(100, self._poll_worker_queue)
        self._show_page("convert")
        self._refresh_home_summary()

    def _worker_run(self, config: BatchConfig) -> None:
        def log(message: str) -> None:
            self.worker_queue.put(("log", message))

        def progress(current: int, total: int) -> None:
            self.worker_queue.put(("progress", (current, total)))

        try:
            outputs = process_batch(config, log=log, progress=progress)
            self.worker_queue.put(("done", (config, outputs)))
        except Exception:
            self.worker_queue.put(("error", (config, traceback.format_exc())))

    def _poll_worker_queue(self) -> None:
        keep_polling = self.running
        try:
            while True:
                message_type, payload = self.worker_queue.get_nowait()
                if message_type == "log":
                    self._append_log(str(payload))
                elif message_type == "progress":
                    current, total = payload  # type: ignore[misc]
                    total = max(int(total), 1)
                    current = int(current)
                    self.progress["value"] = round((current / total) * 100, 2)
                    self.status_var.set(f"Processing {current}/{total}...")
                elif message_type == "done":
                    config, outputs = payload  # type: ignore[assignment]
                    self.running = False
                    self.active_run_kind = ""
                    keep_polling = False
                    self.start_button.state(["!disabled"])
                    self.pdf_tool_start_button.state(["!disabled"])
                    self.progress["value"] = 100
                    self.status_var.set("Completed.")
                    self._track_success_outputs(outputs, output_dir=Path(str(config.output_dir)), label=str(config.mode))
                    self._append_log("Completed successfully.")
                    for item in outputs:
                        self._append_log(f"Created: {item}")
                    self._record_job(config, outputs=outputs, status="Completed")
                    self._refresh_home_summary()
                    self._maybe_auto_open_output(Path(str(config.output_dir)))
                    messagebox.showinfo(
                        "Done",
                        f"Conversion finished successfully.\n\nOutput folder:\n{self.output_dir_var.get()}",
                    )
                elif message_type == "error":
                    config, details = payload  # type: ignore[assignment]
                    self.running = False
                    self.active_run_kind = ""
                    keep_polling = False
                    self.start_button.state(["!disabled"])
                    self.pdf_tool_start_button.state(["!disabled"])
                    self.status_var.set("Failed.")
                    self._append_log("ERROR:\n" + str(details))
                    self._record_job(config, outputs=[], status="Failed", error_text=str(details))
                    self._refresh_home_summary()
                    messagebox.showerror("Conversion failed", str(details))
                elif message_type == "link_status":
                    record = payload  # type: ignore[assignment]
                    self._upsert_link_status(dict(record))
                    status_value = str(record.get("status", ""))
                    if status_value in {"downloaded", "failed", "duplicate", "cancelled", "invalid", "paused", "resumed"}:
                        detail = str(record.get("detail", "") or record.get("error", ""))
                        self._append_log(f"[Links] {status_value.upper()}: {record.get('url', '')} {detail}".rstrip())
                    if status_value == "paused":
                        self.status_var.set("Link downloads paused.")
                    elif status_value == "resumed":
                        self.status_var.set("Link downloads resumed.")
                elif message_type == "link_progress":
                    current, total = payload  # type: ignore[misc]
                    total = max(int(total), 1)
                    current = int(current)
                    self.progress["value"] = round((current / total) * 100, 2)
                    self.status_var.set(f"Downloading links {current}/{total}...")
                elif message_type == "link_done":
                    urls, results, auto_start = payload  # type: ignore[assignment]
                    self.running = False
                    self.active_run_kind = ""
                    keep_polling = False
                    self.start_button.state(["!disabled"])
                    self.pdf_tool_start_button.state(["!disabled"])
                    downloaded = []
                    remembered = []
                    failed = 0
                    for result in results:
                        status_value = getattr(result, "status", "")
                        if status_value == "downloaded" and getattr(result, "file_path", ""):
                            downloaded.append(Path(result.file_path))
                            remembered.append(result.url)
                        elif status_value in {"failed", "invalid"}:
                            failed += 1
                    if remembered:
                        self.state_store.remember_links(remembered)
                    if downloaded:
                        self.link_downloaded_files.extend(downloaded)
                        self._append_files(downloaded)
                    self.progress["value"] = 100
                    if downloaded:
                        self.status_var.set(f"Fetched {len(downloaded)} link file(s).")
                        self.link_status_summary_var.set(
                            f"Fetched {len(downloaded)} file(s) into the local queue. You can convert them now or keep adding more links."
                        )
                    elif self.link_cancel_event.is_set():
                        self.status_var.set("Link download cancelled.")
                    else:
                        self.status_var.set("No files were downloaded from the provided links.")
                    self.link_cancel_event.clear()
                    self.link_pause_event.clear()
                    self._update_link_fetch_buttons()
                    self._refresh_recent_links_summary()
                    self._refresh_link_cache_summary()
                    if downloaded and auto_start:
                        self.after(120, self._start_conversion)
                    elif failed and not downloaded:
                        messagebox.showerror("Link download failed", "None of the provided links could be downloaded. Check the status table for details.")
                elif message_type == "link_error":
                    details = payload
                    self.running = False
                    self.active_run_kind = ""
                    keep_polling = False
                    self.start_button.state(["!disabled"])
                    self.pdf_tool_start_button.state(["!disabled"])
                    self.status_var.set("Link download failed.")
                    self._append_log("LINK ERROR:\n" + str(details))
                    self.link_cancel_event.clear()
                    self.link_pause_event.clear()
                    self._update_link_fetch_buttons()
                    messagebox.showerror("Link download failed", str(details))
                elif message_type == "pdf_log":
                    self._append_pdf_tool_log(str(payload))
                elif message_type == "pdf_progress":
                    current, total = payload  # type: ignore[misc]
                    total = max(int(total), 1)
                    current = int(current)
                    self.pdf_tool_progress["value"] = round((current / total) * 100, 2)
                    self.status_var.set(f"Running PDF tool {current}/{total}...")
                elif message_type == "pdf_done":
                    config, outputs = payload  # type: ignore[assignment]
                    self.running = False
                    self.active_run_kind = ""
                    keep_polling = False
                    self.start_button.state(["!disabled"])
                    self.pdf_tool_start_button.state(["!disabled"])
                    self.pdf_tool_progress["value"] = 100
                    self.status_var.set("PDF tool completed.")
                    self._track_success_outputs(outputs, output_dir=Path(str(config.output_dir)), label=f"PDF Tool -> {config.tool}")
                    self._append_pdf_tool_log("Completed successfully.")
                    self._record_job(config, outputs=outputs, status="Completed")
                    self._maybe_auto_open_output(Path(str(config.output_dir)))
                    messagebox.showinfo(
                        "Done",
                        f"PDF tool finished successfully.\n\nOutput folder:\n{self.output_dir_var.get()}",
                    )
                elif message_type == "pdf_error":
                    config, details = payload  # type: ignore[assignment]
                    self.running = False
                    self.active_run_kind = ""
                    keep_polling = False
                    self.start_button.state(["!disabled"])
                    self.pdf_tool_start_button.state(["!disabled"])
                    self.status_var.set("PDF tool failed.")
                    self._append_pdf_tool_log("ERROR:\n" + str(details))
                    self._record_job(config, outputs=[], status="Failed", error_text=str(details))
                    messagebox.showerror("PDF tool failed", str(details))
        except queue.Empty:
            pass

        if keep_polling:
            self.after(100, self._poll_worker_queue)

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def _record_job(
        self,
        config: BatchConfig | PdfToolConfig,
        outputs: list[Path],
        status: str,
        error_text: str | None = None,
    ) -> None:
        preview = [str(path) for path in config.files[:8]]
        if len(config.files) > 8:
            preview.append(f"... and {len(config.files) - 8} more")

        if isinstance(config, PdfToolConfig):
            record = {
                "job_type": "pdf_tool",
                "status": status,
                "mode": f"PDF Tool -> {config.tool}",
                "tool": config.tool,
                "file_count": len(config.files),
                "output_count": len(outputs),
                "output_dir": str(config.output_dir),
                "output_name": config.output_name,
                "page_spec": config.page_spec,
                "every_n_pages": config.every_n_pages,
                "watermark_text": config.watermark_text,
                "watermark_image": str(config.watermark_image) if config.watermark_image else "",
                "watermark_font_size": config.watermark_font_size,
                "watermark_rotation": config.watermark_rotation,
                "watermark_opacity": config.watermark_opacity,
                "watermark_position": config.watermark_position,
                "watermark_image_scale_percent": config.watermark_image_scale_percent,
                "metadata_title": config.metadata_title,
                "metadata_author": config.metadata_author,
                "metadata_subject": config.metadata_subject,
                "metadata_keywords": config.metadata_keywords,
                "metadata_clear_existing": config.metadata_clear_existing,
                "compression_profile": config.compression_profile,
                "password_used": bool(config.pdf_password),
                "owner_password_used": bool(config.pdf_owner_password),
                "redact_rect": config.redact_rect,
                "replacement_text": config.replacement_text,
                "input_files": [str(path) for path in config.files],
                "inputs_preview": preview,
                "outputs_preview": [str(path) for path in outputs[:8]],
                "error": error_text or "",
            }
        else:
            record = {
                "job_type": "convert",
                "status": status,
                "mode": config.mode,
                "file_count": len(config.files),
                "output_count": len(outputs),
                "output_dir": str(config.output_dir),
                "merge_to_one_pdf": config.merge_to_one_pdf,
                "merged_output_name": config.merged_output_name,
                "image_format": config.image_format,
                "image_scale": config.image_scale,
                "engine_mode": config.engine_mode,
                "soffice_configured": bool(config.soffice_path),
                "input_files": [str(path) for path in config.files],
                "inputs_preview": preview,
                "outputs_preview": [str(path) for path in outputs[:8]],
                "error": error_text or "",
            }
        self.state_store.add_recent_job(record)
        self._refresh_history_views()

    def _refresh_history_views(self) -> None:
        jobs = self.state_store.recent_jobs()
        self.home_history_item_ids.clear()
        self.history_item_ids.clear()

        for tree in (self.home_history_tree, self.history_tree):
            for item in tree.get_children():
                tree.delete(item)

        for index, job in enumerate(jobs):
            tag = "success" if job.get("status") == "Completed" else "error"
            item_id = f"job_{index}"
            values_home = (
                job.get("timestamp", ""),
                job.get("status", ""),
                job.get("mode", ""),
                job.get("file_count", ""),
            )
            values_full = values_home + (job.get("output_count", ""),)
            self.home_history_tree.insert("", tk.END, iid=f"home_{item_id}", values=values_home, tags=(tag,))
            self.history_tree.insert("", tk.END, iid=f"hist_{item_id}", values=values_full, tags=(tag,))
            self.home_history_item_ids[f"home_{item_id}"] = job
            self.history_item_ids[f"hist_{item_id}"] = job

        self.history_detail_var.set("Select a recent job to inspect details or re-use settings.")
        self._set_history_details_text("")
        self._refresh_recent_outputs_view()
        self._refresh_failed_jobs_view()

    def _on_home_history_selected(self, _event=None) -> None:
        item = self.home_history_tree.focus()
        job = self.home_history_item_ids.get(item)
        if not job:
            return
        self.history_detail_var.set(self._job_summary_text(job))

    def _on_history_selected(self, _event=None) -> None:
        item = self.history_tree.focus()
        job = self.history_item_ids.get(item)
        if not job:
            return
        self.history_detail_var.set(self._job_summary_text(job))
        self._set_history_details_text(self._job_detail_text(job))

    def _job_summary_text(self, job: dict[str, object]) -> str:
        return (
            f"{job.get('status', '')}  |  {job.get('mode', '')}\n"
            f"Files: {job.get('file_count', 0)}  |  Outputs: {job.get('output_count', 0)}\n"
            f"Output folder: {job.get('output_dir', '')}"
        )

    def _job_detail_text(self, job: dict[str, object]) -> str:
        inputs_preview = job.get("inputs_preview", []) or []
        outputs_preview = job.get("outputs_preview", []) or []
        full_inputs = job.get("input_files", []) or []
        lines = [
            f"Timestamp: {job.get('timestamp', '')}",
            f"Status: {job.get('status', '')}",
            f"Mode: {job.get('mode', '')}",
            f"Files selected: {job.get('file_count', 0)}",
            f"Outputs created: {job.get('output_count', 0)}",
            f"Output folder: {job.get('output_dir', '')}",
        ]

        job_type = str(job.get("job_type", "convert"))
        if job_type == "pdf_tool":
            lines.extend(
                [
                    f"PDF tool: {job.get('tool', '')}",
                    f"Merged output name: {job.get('output_name', '')}",
                    f"Page spec: {job.get('page_spec', '')}",
                    f"Split every N pages: {job.get('every_n_pages', '')}",
                    f"Watermark text: {job.get('watermark_text', '')}",
                    f"Watermark image: {job.get('watermark_image', '')}",
                    f"Watermark font size: {job.get('watermark_font_size', '')}",
                    f"Watermark rotation: {job.get('watermark_rotation', '')}",
                    f"Watermark opacity: {job.get('watermark_opacity', '')}",
                    f"Watermark position: {job.get('watermark_position', '')}",
                    f"Watermark image scale %: {job.get('watermark_image_scale_percent', '')}",
                    f"Metadata title: {job.get('metadata_title', '')}",
                    f"Metadata author: {job.get('metadata_author', '')}",
                    f"Metadata subject: {job.get('metadata_subject', '')}",
                    f"Metadata keywords: {job.get('metadata_keywords', '')}",
                    f"Metadata clear existing: {job.get('metadata_clear_existing', False)}",
                    f"Compression profile: {job.get('compression_profile', '')}",
                    f"Redact rectangle: {job.get('redact_rect', '')}",
                    f"Replacement text: {job.get('replacement_text', '')}",
                    f"Password supplied for run: {job.get('password_used', False)}",
                    f"Owner password supplied: {job.get('owner_password_used', False)}",
                ]
            )
        elif job_type == "ocr":
            lines.extend(
                [
                    f"OCR language: {job.get('ocr_language', '')}",
                    f"OCR DPI: {job.get('ocr_dpi', '')}",
                    f"OCR PSM: {job.get('ocr_psm', '')}",
                    f"Tesseract configured for run: {job.get('tesseract_configured', False)}",
                ]
            )
        elif job_type == "organizer":
            lines.extend(
                [
                    f"Organizer note: {job.get('note', '')}",
                ]
            )
        else:
            lines.extend(
                [
                    f"Merge into one PDF: {job.get('merge_to_one_pdf', False)}",
                    f"Merged output name: {job.get('merged_output_name', '')}",
                    f"Image format: {job.get('image_format', '')}",
                    f"Image scale: {job.get('image_scale', '')}",
                    f"Engine mode: {job.get('engine_mode', ENGINE_AUTO)}",
                    f"Soffice configured for run: {job.get('soffice_configured', False)}",
                ]
            )

        lines.extend(
            [
                "",
                "Inputs preview:",
                *[f"- {line}" for line in inputs_preview],
                "",
                "Outputs preview:",
                *[f"- {line}" for line in outputs_preview],
            ]
        )
        error_text = str(job.get("error", "")).strip()
        if error_text:
            lines.extend(["", "Error details:", error_text])
        return "\n".join(lines).strip() + "\n"

    def _set_history_details_text(self, text: str) -> None:
        self.history_details_text.configure(state="normal")
        self.history_details_text.delete("1.0", tk.END)
        self.history_details_text.insert(tk.END, text)
        self.history_details_text.configure(state="disabled")

    def _load_selected_history_settings(self) -> None:
        item = self.history_tree.focus()
        job = self.history_item_ids.get(item)
        if not job:
            messagebox.showinfo("History", "Please select a job in the history list first.")
            return

        self.output_dir_var.set(str(job.get("output_dir", self.output_dir_var.get())))
        job_type = str(job.get("job_type", "convert"))

        if job_type == "pdf_tool":
            self.pdf_tool_var.set(str(job.get("tool", PDF_TOOL_MERGE)))
            self.pdf_tool_output_name_var.set(str(job.get("output_name", "merged_pdfs")))
            self.pdf_tool_page_spec_var.set(str(job.get("page_spec", "")))
            self.pdf_tool_every_n_var.set(str(job.get("every_n_pages", "2")))
            self.pdf_tool_watermark_text_var.set(str(job.get("watermark_text", "CONFIDENTIAL")))
            self.pdf_tool_watermark_image_var.set(str(job.get("watermark_image", "")))
            self.pdf_tool_font_size_var.set(str(job.get("watermark_font_size", "42")))
            self.pdf_tool_rotation_var.set(str(job.get("watermark_rotation", "45")))
            self.pdf_tool_opacity_var.set(str(job.get("watermark_opacity", "0.18")))
            self.pdf_tool_position_var.set(str(job.get("watermark_position", "center")))
            self.pdf_tool_image_scale_var.set(str(job.get("watermark_image_scale_percent", "40")))
            self.pdf_tool_metadata_title_var.set(str(job.get("metadata_title", "")))
            self.pdf_tool_metadata_author_var.set(str(job.get("metadata_author", "")))
            self.pdf_tool_metadata_subject_var.set(str(job.get("metadata_subject", "")))
            self.pdf_tool_metadata_keywords_var.set(str(job.get("metadata_keywords", "")))
            self.pdf_tool_metadata_clear_var.set(bool(job.get("metadata_clear_existing", False)))
            self.pdf_tool_redact_rect_var.set(str(job.get("redact_rect", self.pdf_tool_redact_rect_var.get())))
            self.pdf_tool_replacement_text_var.set(str(job.get("replacement_text", "")))
            self.pdf_tool_password_var.set("")
            self.pdf_tool_owner_password_var.set("")
            self.pdf_tool_compression_profile_var.set(str(job.get("compression_profile", "balanced")))
            self._update_pdf_tool_controls()
            self._show_page("pdf_tools")
        elif job_type == "ocr":
            self.ocr_output_var.set(str(job.get("output_dir", self.ocr_output_var.get())))
            self.ocr_language_var.set(str(job.get("ocr_language", self.ocr_language_var.get())))
            self.ocr_dpi_var.set(str(job.get("ocr_dpi", self.ocr_dpi_var.get())))
            self.ocr_psm_var.set(str(job.get("ocr_psm", self.ocr_psm_var.get())))
            self._refresh_dependency_status()
            self._show_page("ocr")
        else:
            self.mode_var.set(str(job.get("mode", MODE_ANY_TO_PDF)))
            self.merge_var.set(bool(job.get("merge_to_one_pdf", False)))
            self.output_name_var.set(str(job.get("merged_output_name", default_merged_name(self.mode_var.get()))))
            self.image_format_var.set(str(job.get("image_format", "png")))
            self.image_scale_var.set(str(job.get("image_scale", "2.0")))
            self.engine_mode_var.set(str(job.get("engine_mode", ENGINE_AUTO)))
            self._refresh_dependency_status()
            self._update_mode_controls()
            self._show_page("convert")

        self.status_var.set("Loaded settings from selected history item.")

    def _open_selected_history_output(self) -> None:
        item = self.history_tree.focus()
        job = self.history_item_ids.get(item)
        if not job:
            messagebox.showinfo("History", "Please select a job in the history list first.")
            return
        open_path(Path(str(job.get("output_dir", self.output_dir_var.get()))))

    def _open_mail_draft_for_last_outputs(self) -> None:
        if not self.last_outputs:
            messagebox.showinfo("Mail draft", "Run a conversion, PDF tool, or OCR task first so the app has outputs to reference.")
            return

        subject, body = self._build_output_email_content()
        recipients = self.smtp_default_to_var.get().strip()
        try:
            open_mailto_draft(recipients, subject=subject, body=body)
            self.status_var.set("Opened a mail draft for the latest outputs.")
        except Exception as exc:
            messagebox.showerror("Mail draft", f"Could not open a mail draft:\n{exc}")

    def _save_eml_draft_for_last_outputs(self) -> None:
        if not self.last_outputs:
            messagebox.showinfo("EML draft", "Run a conversion, PDF tool, or OCR task first so the app has outputs to package.")
            return

        output_dir = self.last_output_dir or Path(self.output_dir_var.get().strip() or Path.cwd() / "converted_output")
        suggested_name = f"{(self.last_job_label or 'latest_outputs').replace('/', '_').replace(':', '_').replace(' ', '_')}.eml"
        target = filedialog.asksaveasfilename(
            title="Save EML draft",
            defaultextension=".eml",
            initialdir=str(output_dir),
            initialfile=suggested_name,
            filetypes=[("EML message", "*.eml"), ("All files", "*.*")],
        )
        if not target:
            return

        sender = self.smtp_sender_var.get().strip() or str(self.about_profile.get("email", "")).strip()
        if not sender:
            messagebox.showerror("EML draft", "Set a sender email in SMTP settings or your About profile first.")
            return

        recipients = self.smtp_default_to_var.get().strip() or "recipient@example.com"
        subject, body = self._build_output_email_content()
        try:
            eml_path = build_eml_draft(
                target,
                sender=sender,
                recipients=recipients,
                subject=subject,
                body=body,
                attachments=self.last_outputs,
            )
        except Exception as exc:
            messagebox.showerror("EML draft", f"Could not create the EML draft:\n{exc}")
            return

        if Path(eml_path) not in self.last_outputs:
            self.last_outputs = [*self.last_outputs, Path(eml_path)]
        self.last_output_dir = output_dir
        self.status_var.set(f"Saved EML draft: {Path(eml_path).name}")
        messagebox.showinfo("EML draft", f"Saved EML draft to:\n{eml_path}")


    def _clear_history(self) -> None:
        if not self.state_store.recent_jobs():
            return
        if not messagebox.askyesno("Clear history", "Remove all recent job history from this machine?"):
            return
        self.state_store.clear_recent_jobs()
        self._refresh_history_views()
        self.status_var.set("History cleared.")

    def _set_theme(self, choice: str) -> None:
        self.theme_choice_var.set(choice)
        self._apply_theme()

    def _apply_theme(self, initial: bool = False) -> None:
        choice = self.theme_choice_var.get().strip().lower() or "dark"
        palette = resolve_palette(choice)
        self.palette = palette
        apply_ttk_theme(self, palette)
        self.option_add("*TCombobox*Listbox.background", palette.input_bg)
        self.option_add("*TCombobox*Listbox.foreground", palette.text)
        self.option_add("*TCombobox*Listbox.selectBackground", palette.selection)
        self.option_add("*TCombobox*Listbox.selectForeground", palette.accent_text)
        for menu in self.menus:
            apply_menu_theme(menu, palette)
        for widget in (self.file_listbox, self.link_input_text, self.log_text, self.pdf_tool_listbox, self.pdf_tool_log_text, self.ocr_log_text, self.history_details_text, self.automation_log_text):
            apply_text_widget_theme(widget, palette)
        if self.organizer_panel is not None:
            self.organizer_panel.apply_theme(palette)
        apply_treeview_tag_colors(self.home_history_tree, palette)
        apply_treeview_tag_colors(self.history_tree, palette)
        if self.notes_window and self.notes_window.winfo_exists():
            self.notes_window.apply_theme(palette)
        if self.smtp_window and self.smtp_window.winfo_exists():
            self.smtp_window.apply_theme(palette)
        if self.build_center_window and self.build_center_window.winfo_exists():
            self.build_center_window.apply_theme(palette)
            self.build_center_window.refresh_summary()
        if self.about_editor_window and self.about_editor_window.winfo_exists():
            self.about_editor_window.apply_theme(palette)
        if self.splash_window and self.splash_window.winfo_exists():
            apply_ttk_theme(self.splash_window, palette)
        if self.login_popup_window and self.login_popup_window.winfo_exists():
            apply_ttk_theme(self.login_popup_window, palette)
        if not initial:
            self.status_var.set(f"Theme changed to {choice}.")

    def _open_notes_window(self) -> None:
        palette = resolve_palette(self.theme_choice_var.get())
        if self.notes_window and self.notes_window.winfo_exists():
            self.notes_window.lift()
            self.notes_window.focus_force()
            self.notes_window.reload_notes()
            return
        self.notes_window = MarkdownNotesWindow(self, self.notes_path, palette)

    def _smtp_settings_from_vars(self) -> SMTPSettings:
        port_text = self.smtp_port_var.get().strip()
        try:
            port = int(port_text) if port_text else 587
        except Exception as exc:
            raise ValueError("SMTP port must be a number.") from exc
        return SMTPSettings(
            host=self.smtp_host_var.get().strip(),
            port=port,
            username=self.smtp_username_var.get().strip(),
            password=self.smtp_password_var.get(),
            sender=self.smtp_sender_var.get().strip(),
            default_to=self.smtp_default_to_var.get().strip(),
            use_ssl=bool(self.smtp_use_ssl_var.get()),
            use_starttls=bool(self.smtp_use_starttls_var.get()),
            save_password=bool(self.smtp_save_password_var.get()),
        )

    def _save_smtp_settings(self) -> None:
        try:
            settings = self._smtp_settings_from_vars()
            settings.validate_for_send()
        except Exception as exc:
            self.smtp_status_var.set(str(exc))
            messagebox.showerror("SMTP settings", str(exc))
            return
        self._persist_state()
        self.smtp_status_var.set("SMTP settings saved locally.")
        self.status_var.set("SMTP settings saved.")

    def _test_smtp_settings(self) -> None:
        try:
            settings = self._smtp_settings_from_vars()
            summary = test_smtp_connection(settings)
        except Exception as exc:
            self.smtp_status_var.set(str(exc))
            messagebox.showerror("SMTP test", f"Could not validate SMTP settings:\n{exc}")
            return
        self.smtp_status_var.set(summary)
        self._persist_state()
        messagebox.showinfo("SMTP test", summary)
        self.status_var.set("SMTP connection test succeeded.")

    def _build_output_email_content(self) -> tuple[str, str]:
        output_dir = self.last_output_dir or Path(self.output_dir_var.get().strip() or Path.cwd() / "converted_output")
        preview_items = self.last_outputs[:12]
        body_lines = [
            "Hello,",
            "",
            f"These files were generated with {APP_NAME}.",
            "",
        ]
        body_lines.extend(f"- {path.name}" for path in preview_items)
        if len(self.last_outputs) > len(preview_items):
            body_lines.append(f"- ... and {len(self.last_outputs) - len(preview_items)} more")
        body_lines.extend([
            "",
            f"Output folder: {output_dir}",
            "",
            "Sent directly from Gokul Omni Convert Lite.",
        ])
        subject = f"{APP_NAME} outputs - {self.last_job_label or 'latest run'}"
        return subject, "\n".join(body_lines)

    def _send_last_outputs_via_smtp(self) -> None:
        if not self.last_outputs:
            messagebox.showinfo("SMTP send", "Run a conversion or PDF tool first so the app has outputs to send.")
            return
        try:
            settings = self._smtp_settings_from_vars()
            settings.validate_for_send()
        except Exception as exc:
            self.smtp_status_var.set(str(exc))
            messagebox.showerror("SMTP send", str(exc))
            return

        default_recipient = settings.default_to
        recipients = simpledialog.askstring(
            "Send outputs",
            "Recipient email address(es), separated by commas:",
            initialvalue=default_recipient,
            parent=self,
        )
        if recipients is None:
            return
        recipients = recipients.strip()
        if not recipients:
            messagebox.showerror("SMTP send", "Enter at least one recipient email address.")
            return

        subject, body = self._build_output_email_content()
        try:
            delivered = send_email(settings, recipients, subject, body, self.last_outputs)
        except Exception as exc:
            self.smtp_status_var.set(str(exc))
            messagebox.showerror("SMTP send", f"Could not send the latest outputs:\n{exc}")
            return

        self.smtp_default_to_var.set(recipients)
        self._persist_state()
        self.smtp_status_var.set(f"Sent {len(delivered)} attachment(s) to {recipients}.")
        self.status_var.set("Latest outputs sent via SMTP.")
        messagebox.showinfo("SMTP send", f"Message sent successfully with {len(delivered)} attachment(s).")

    def _open_smtp_window(self) -> None:
        palette = resolve_palette(self.theme_choice_var.get())
        if self.smtp_window and self.smtp_window.winfo_exists():
            self.smtp_window.lift()
            self.smtp_window.focus_force()
            return
        self.smtp_window = SMTPSettingsWindow(self, palette)

    def _open_build_center_window(self) -> None:
        palette = resolve_palette(self.theme_choice_var.get())
        if self.build_center_window and self.build_center_window.winfo_exists():
            self.build_center_window.lift()
            self.build_center_window.focus_force()
            self.build_center_window.refresh_summary()
            return
        self.build_center_window = BuildCenterWindow(self, palette)

    def _open_about_editor_window(self) -> None:
        palette = resolve_palette(self.theme_choice_var.get())
        if self.about_editor_window and self.about_editor_window.winfo_exists():
            self.about_editor_window.lift()
            self.about_editor_window.focus_force()
            self.about_editor_window.load_profile()
            return
        self.about_editor_window = AboutProfileEditorWindow(self, palette)

    def _open_installer_folder(self) -> None:
        open_path(self.build_notes_path.parent)

    def _export_state_snapshot_action(self) -> None:
        self._persist_state()
        target = filedialog.asksaveasfilename(
            title="Export settings snapshot",
            defaultextension=".json",
            initialfile="gokul_omni_convert_lite_settings.json",
            filetypes=[("JSON", "*.json")],
        )
        if not target:
            return
        destination = export_state_snapshot(Path(target), self.state_store.state)
        self.status_var.set(f"Settings snapshot exported to {destination.name}.")
        messagebox.showinfo("Settings snapshot", f"Exported settings snapshot to:\n{destination}")

    def _apply_state_snapshot(self, payload: dict[str, object]) -> None:
        self.state_store.state = {**self.state_store.state, **payload}
        ensure_install_date(self.state_store.state)
        self.state_store.save()
        self.theme_choice_var.set(str(self.state_store.get("theme", "dark")))
        self.output_dir_var.set(str(self.state_store.get("output_dir", str(Path.cwd() / "converted_output"))))
        self.recursive_var.set(bool(self.state_store.get("recursive_scan", True)))
        self.engine_mode_var.set(str(self.state_store.get("conversion_engine", ENGINE_AUTO)))
        self.soffice_path_var.set(str(self.state_store.get("soffice_path", "")))
        self.tesseract_path_var.set(str(self.state_store.get("tesseract_path", "")))
        self.ocr_language_var.set(str(self.state_store.get("ocr_language", "eng")))
        self.ocr_dpi_var.set(str(self.state_store.get("ocr_dpi", 220)))
        self.ocr_psm_var.set(str(self.state_store.get("ocr_psm", 6)))
        self.ocr_output_var.set(str(self.state_store.get("ocr_output_dir", "")))
        self.splash_enabled_var.set(bool(self.state_store.get("splash_enabled", True)))
        self.splash_gif_path_var.set(str(self.state_store.get("splash_gif_path", "assets/gokul_splash.gif")))
        self.login_popup_enabled_var.set(bool(self.state_store.get("login_popup_enabled", True)))
        self.install_date_var.set(str(self.state_store.get("install_date", "")))
        smtp_config = SMTPSettings.from_dict(self.state_store.get("smtp_settings", {}))
        self.smtp_host_var.set(smtp_config.host)
        self.smtp_port_var.set(str(smtp_config.port))
        self.smtp_username_var.set(smtp_config.username)
        self.smtp_password_var.set(smtp_config.password)
        self.smtp_sender_var.set(smtp_config.sender or str(self.about_profile.get("email", "")).strip())
        self.smtp_default_to_var.set(smtp_config.default_to)
        self.smtp_use_ssl_var.set(smtp_config.use_ssl)
        self.smtp_use_starttls_var.set(smtp_config.use_starttls)
        self.smtp_save_password_var.set(smtp_config.save_password)
        watch_config = normalize_watch_config(self.state_store.watch_config())
        self.watch_source_dir_var.set(watch_config.source_dir)
        self.watch_output_dir_var.set(watch_config.output_dir or self.output_dir_var.get().strip())
        self.watch_mode_var.set(watch_config.mode)
        self.watch_merge_var.set(watch_config.merge_to_one_pdf)
        self.watch_recursive_var.set(watch_config.recursive)
        self.watch_interval_var.set(str(watch_config.interval_seconds))
        self.watch_output_name_var.set(watch_config.merged_output_name)
        self.watch_engine_var.set(watch_config.engine_mode)
        self.watch_archive_var.set(watch_config.archive_processed)
        self.watch_archive_dir_var.set(watch_config.archive_dir)
        self.watch_zip_var.set(watch_config.create_zip_bundle)
        self.watch_report_var.set(watch_config.create_report)
        self.watch_mail_var.set(watch_config.open_mail_draft)
        self.watch_skip_existing_var.set(watch_config.skip_existing_on_start)
        self._update_login_popup_state()
        self._refresh_dependency_status()
        self._refresh_about_profile()
        self._refresh_history_views()
        self._apply_theme()

    def _import_state_snapshot_action(self) -> None:
        source = filedialog.askopenfilename(
            title="Import settings snapshot",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not source:
            return
        try:
            payload = import_state_snapshot(Path(source))
        except Exception as exc:
            messagebox.showerror("Import settings", f"Could not import the selected snapshot:\n{exc}")
            return
        self._apply_state_snapshot(payload)
        self.status_var.set("Settings snapshot imported.")
        messagebox.showinfo("Import settings", "Settings snapshot imported successfully.")

    def _export_diagnostics_report_action(self) -> None:
        self._persist_state()
        target = filedialog.asksaveasfilename(
            title="Export diagnostics report",
            defaultextension=".json",
            initialfile="gokul_omni_convert_lite_diagnostics.json",
            filetypes=[("JSON", "*.json")],
        )
        if not target:
            return
        runtime_dependencies = dependency_status(self.soffice_path_var.get().strip())
        tesseract_runtime = detect_tesseract_status(self.tesseract_path_var.get().strip())
        runtime_dependencies["Tesseract"] = bool(tesseract_runtime.get("available"))
        destination = export_diagnostics_report(
            Path(target),
            app_name=APP_NAME,
            app_version=APP_VERSION,
            state_path=APP_STATE_PATH,
            about_profile_path=self.about_profile_path,
            notes_path=self.notes_path,
            installer_dir=self.build_notes_path.parent,
            output_dir=Path(self.output_dir_var.get().strip() or Path.cwd() / "converted_output"),
            selected_files=[str(path) for path in self.selected_files],
            last_outputs=[str(path) for path in self.last_outputs],
            dependency_status=runtime_dependencies,
            smtp_summary=self._smtp_settings_from_vars().sanitized_dict(),
            extra={
                "recent_job_count": len(self.state_store.recent_jobs()),
                "preset_count": len(self.state_store.presets()),
                "last_job_label": self.last_job_label,
                "last_run_origin": self.last_run_origin,
                "tesseract": {
                    "available": bool(tesseract_runtime.get("available")),
                    "source": str(tesseract_runtime.get("source", "")),
                    "path": str(tesseract_runtime.get("path", "")),
                },
                "ocr_defaults": {
                    "language": self.ocr_language_var.get().strip(),
                    "dpi": self.ocr_dpi_var.get().strip(),
                    "psm": self.ocr_psm_var.get().strip(),
                },
            },
        )
        self.status_var.set(f"Diagnostics report exported to {destination.name}.")
        messagebox.showinfo("Diagnostics report", f"Exported diagnostics report to:\n{destination}")

    def _show_about(self) -> None:
        self._refresh_about_profile()
        self._show_page("about")
        self.status_var.set("About page opened.")


    def _build_automation_page(self) -> None:
        page = ttk.Frame(self.content)
        page.grid(row=0, column=0, sticky="nsew")
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(1, weight=1)
        self.pages["automation"] = page
        self.preset_item_ids: dict[str, dict[str, object]] = {}

        header = ttk.Frame(page, style="Card.TFrame", padding=22)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        ttk.Label(header, text="Automation, presets, and output sharing", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text=(
                "Automation adds reusable presets, a watch-folder workflow, and built-in ways to bundle or report the latest outputs. "
                "The watcher uses your current built-in conversion engine choices and can optionally archive processed source files."
            ),
            style="CardBody.TLabel",
            wraplength=980,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        content = ttk.Frame(page)
        content.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        content.grid_columnconfigure(0, weight=1)
        content.grid_columnconfigure(1, weight=1)
        content.grid_rowconfigure(0, weight=1)
        content.grid_rowconfigure(1, weight=1)

        presets_frame = ttk.LabelFrame(content, text="Saved presets")
        presets_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=(0, 10))
        presets_frame.grid_columnconfigure(0, weight=1)
        presets_frame.grid_rowconfigure(2, weight=1)

        ttk.Label(
            presets_frame,
            text="Save the current Convert screen settings as a named preset, then re-apply them later from here.",
            style="CardBody.TLabel",
            wraplength=460,
            justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

        name_row = ttk.Frame(presets_frame)
        name_row.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        name_row.grid_columnconfigure(1, weight=1)
        ttk.Label(name_row, text="Preset name:", style="Surface.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(name_row, textvariable=self.preset_name_var).grid(row=0, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(name_row, text="Save current", command=self._save_current_preset).grid(row=0, column=2, sticky="e")

        self.preset_tree = ttk.Treeview(
            presets_frame,
            columns=("favorite", "name", "mode", "engine", "merge"),
            show="headings",
            height=9,
        )
        self.preset_tree.heading("favorite", text="★")
        self.preset_tree.heading("name", text="Name")
        self.preset_tree.heading("mode", text="Mode")
        self.preset_tree.heading("engine", text="Engine")
        self.preset_tree.heading("merge", text="Merge")
        self.preset_tree.column("favorite", width=44, anchor="center")
        self.preset_tree.column("name", width=140, anchor="w")
        self.preset_tree.column("mode", width=190, anchor="w")
        self.preset_tree.column("engine", width=90, anchor="center")
        self.preset_tree.column("merge", width=70, anchor="center")
        self.preset_tree.grid(row=2, column=0, columnspan=3, sticky="nsew")
        self.preset_tree.bind("<<TreeviewSelect>>", self._on_preset_selected)

        preset_actions = ttk.Frame(presets_frame)
        preset_actions.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        for index in range(5):
            preset_actions.grid_columnconfigure(index, weight=1)
        ttk.Button(preset_actions, text="Apply to Convert", command=self._apply_selected_preset).grid(row=0, column=0, sticky="ew")
        ttk.Button(preset_actions, text="★ Favorite", command=self._toggle_selected_preset_favorite).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        ttk.Button(preset_actions, text="Delete", command=self._delete_selected_preset).grid(row=0, column=2, sticky="ew", padx=(8, 0))
        ttk.Button(preset_actions, text="Export JSON", command=self._export_presets).grid(row=0, column=3, sticky="ew", padx=(8, 0))
        ttk.Button(preset_actions, text="Import JSON", command=self._import_presets).grid(row=0, column=4, sticky="ew", padx=(8, 0))

        watch_frame = ttk.LabelFrame(content, text="Watch folder automation")
        watch_frame.grid(row=0, column=1, sticky="nsew", pady=(0, 10))
        watch_frame.grid_columnconfigure(1, weight=1)
        watch_frame.grid_columnconfigure(2, weight=0)

        ttk.Label(watch_frame, text="Source folder:", style="Surface.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(watch_frame, textvariable=self.watch_source_dir_var).grid(row=0, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(watch_frame, text="Browse", command=self._browse_watch_source_dir).grid(row=0, column=2, sticky="e")

        ttk.Label(watch_frame, text="Output folder:", style="Surface.TLabel").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(watch_frame, textvariable=self.watch_output_dir_var).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=(8, 0))
        ttk.Button(watch_frame, text="Browse", command=self._browse_watch_output_dir).grid(row=1, column=2, sticky="e", pady=(8, 0))

        ttk.Label(watch_frame, text="Mode:", style="Surface.TLabel").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Combobox(watch_frame, textvariable=self.watch_mode_var, values=MODE_ORDER, state="readonly").grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=(8, 0))
        ttk.Label(watch_frame, text="Engine:", style="Surface.TLabel").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Combobox(watch_frame, textvariable=self.watch_engine_var, values=ENGINE_ORDER, state="readonly").grid(row=3, column=1, sticky="ew", padx=(8, 8), pady=(8, 0))

        ttk.Checkbutton(watch_frame, text="Merge outputs into one PDF when supported", variable=self.watch_merge_var).grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(10, 0)
        )
        merged_row = ttk.Frame(watch_frame)
        merged_row.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        merged_row.grid_columnconfigure(1, weight=1)
        ttk.Label(merged_row, text="Merged output name:", style="Surface.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(merged_row, textvariable=self.watch_output_name_var).grid(row=0, column=1, sticky="ew", padx=(8, 0))

        interval_row = ttk.Frame(watch_frame)
        interval_row.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Checkbutton(interval_row, text="Recursive scan", variable=self.watch_recursive_var).grid(row=0, column=0, sticky="w")
        ttk.Label(interval_row, text="Interval seconds:", style="Surface.TLabel").grid(row=0, column=1, sticky="e", padx=(18, 8))
        ttk.Entry(interval_row, width=8, textvariable=self.watch_interval_var).grid(row=0, column=2, sticky="w")

        ttk.Checkbutton(watch_frame, text="Archive processed source files", variable=self.watch_archive_var).grid(
            row=7, column=0, columnspan=3, sticky="w", pady=(10, 0)
        )
        archive_row = ttk.Frame(watch_frame)
        archive_row.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        archive_row.grid_columnconfigure(1, weight=1)
        ttk.Label(archive_row, text="Archive folder:", style="Surface.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(archive_row, textvariable=self.watch_archive_dir_var).grid(row=0, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(archive_row, text="Browse", command=self._browse_watch_archive_dir).grid(row=0, column=2, sticky="e")

        ttk.Checkbutton(watch_frame, text="Create ZIP bundle after each automation run", variable=self.watch_zip_var).grid(
            row=9, column=0, columnspan=3, sticky="w", pady=(10, 0)
        )
        ttk.Checkbutton(watch_frame, text="Create text report after each automation run", variable=self.watch_report_var).grid(
            row=10, column=0, columnspan=3, sticky="w", pady=(4, 0)
        )
        ttk.Checkbutton(watch_frame, text="Open a mail draft after each automation run", variable=self.watch_mail_var).grid(
            row=11, column=0, columnspan=3, sticky="w", pady=(4, 0)
        )
        ttk.Checkbutton(watch_frame, text="Skip existing files when watcher starts", variable=self.watch_skip_existing_var).grid(
            row=12, column=0, columnspan=3, sticky="w", pady=(4, 0)
        )

        watch_actions = ttk.Frame(watch_frame)
        watch_actions.grid(row=13, column=0, columnspan=3, sticky="ew", pady=(14, 0))
        for index in range(4):
            watch_actions.grid_columnconfigure(index, weight=1)
        ttk.Button(watch_actions, text="Save config", command=self._save_watch_config).grid(row=0, column=0, sticky="ew")
        ttk.Button(watch_actions, text="Scan now", command=self._scan_watch_folder_now).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        ttk.Button(watch_actions, text="Start watcher", command=self._start_watch_automation).grid(row=0, column=2, sticky="ew", padx=(8, 0))
        ttk.Button(watch_actions, text="Stop watcher", command=self._stop_watch_automation).grid(row=0, column=3, sticky="ew", padx=(8, 0))

        ttk.Label(watch_frame, textvariable=self.watch_status_var, style="CardBody.TLabel", wraplength=460, justify="left").grid(
            row=14, column=0, columnspan=3, sticky="w", pady=(12, 0)
        )

        lower = ttk.Frame(content)
        lower.grid(row=1, column=0, columnspan=2, sticky="nsew")
        lower.grid_columnconfigure(0, weight=1)
        lower.grid_columnconfigure(1, weight=1)
        lower.grid_rowconfigure(0, weight=1)

        share_frame = ttk.LabelFrame(lower, text="Bundles and reports")
        share_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        share_frame.grid_columnconfigure(1, weight=1)
        ttk.Label(share_frame, textvariable=self.watch_summary_var, style="CardBody.TLabel", wraplength=460, justify="left").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 12)
        )

        ttk.Label(share_frame, text="Bundle base name:", style="Surface.TLabel").grid(row=1, column=0, sticky="w")
        ttk.Entry(share_frame, textvariable=self.share_bundle_name_var).grid(row=1, column=1, columnspan=2, sticky="ew", padx=(8, 0))
        ttk.Label(share_frame, text="Report base name:", style="Surface.TLabel").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(share_frame, textvariable=self.share_report_name_var).grid(row=2, column=1, columnspan=2, sticky="ew", padx=(8, 0), pady=(8, 0))

        bundle_actions = ttk.Frame(share_frame)
        bundle_actions.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(14, 0))
        for index in range(4):
            bundle_actions.grid_columnconfigure(index, weight=1)
        ttk.Button(bundle_actions, text="ZIP last outputs", command=self._create_zip_bundle_for_last_outputs).grid(row=0, column=0, sticky="ew")
        ttk.Button(bundle_actions, text="Export last report", command=self._export_last_run_report).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        ttk.Button(bundle_actions, text="Mail draft", command=self._open_mail_draft_for_last_outputs).grid(row=0, column=2, sticky="ew", padx=(8, 0))
        ttk.Button(bundle_actions, text="Save EML", command=self._save_eml_draft_for_last_outputs).grid(row=0, column=3, sticky="ew", padx=(8, 0))

        utility_actions = ttk.Frame(share_frame)
        utility_actions.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        for index in range(4):
            utility_actions.grid_columnconfigure(index, weight=1)
        ttk.Button(utility_actions, text="Open last output folder", command=self._open_last_output_folder).grid(row=0, column=0, sticky="ew")
        ttk.Button(utility_actions, text="Send latest outputs", command=self._send_last_outputs_via_smtp).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        ttk.Button(utility_actions, text="Reset seen cache", command=self._clear_watch_seen_cache).grid(row=0, column=2, sticky="ew", padx=(8, 0))
        ttk.Button(utility_actions, text="Refresh page", command=self._refresh_automation_page).grid(row=0, column=3, sticky="ew", padx=(8, 0))

        log_frame = ttk.LabelFrame(lower, text="Automation event log")
        log_frame.grid(row=0, column=1, sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)
        self.automation_log_text = ScrolledText(log_frame, wrap=tk.WORD, relief="flat", borderwidth=1, padx=12, pady=10, height=16)
        self.automation_log_text.grid(row=0, column=0, sticky="nsew")
        self.automation_log_text.configure(state="disabled")

        self._restore_automation_log()
        self._refresh_presets_view()
        self._refresh_automation_page()

    def _restore_automation_log(self) -> None:
        if not hasattr(self, "automation_log_text"):
            return
        self.automation_log_text.configure(state="normal")
        self.automation_log_text.delete("1.0", tk.END)
        events = list(reversed(self.state_store.automation_events()))
        if not events:
            self.automation_log_text.insert(tk.END, "No automation events yet.\n")
        else:
            for event in events:
                timestamp = str(event.get("timestamp", ""))
                message = str(event.get("message", "")).strip()
                if not message:
                    continue
                self.automation_log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.automation_log_text.see(tk.END)
        self.automation_log_text.configure(state="disabled")

    def _log_automation(self, message: str, persist: bool = True) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{timestamp}] {message}".rstrip()
        if hasattr(self, "automation_log_text"):
            self.automation_log_text.configure(state="normal")
            self.automation_log_text.insert(tk.END, line + "\n")
            self.automation_log_text.see(tk.END)
            self.automation_log_text.configure(state="disabled")
        if persist:
            self.state_store.add_automation_event({"message": message})
        self.status_var.set(message)

    def _refresh_automation_page(self) -> None:
        preset_count = len(self.state_store.presets())
        seen_count = len(self.state_store.watch_seen_fingerprints())
        watcher_state = "running" if self.automation_active else "idle"
        source_label = Path(self.watch_source_dir_var.get().strip()).name if self.watch_source_dir_var.get().strip() else "no watch folder"
        self.watch_summary_var.set(
            f"Presets saved: {preset_count}. Cached processed files: {seen_count}. Watcher state: {watcher_state} ({source_label}). "
            f"Use ZIP/report buttons to package the latest outputs for sharing."
        )

    def _selected_preset(self) -> dict[str, object] | None:
        item = self.preset_tree.focus() if hasattr(self, "preset_tree") else ""
        return self.preset_item_ids.get(item)

    def _on_preset_selected(self, _event=None) -> None:
        preset = self._selected_preset()
        if not preset:
            return
        self.preset_name_var.set(str(preset.get("name", "My preset")))
        star = "★ " if bool(preset.get("favorite", False)) else ""
        self.watch_summary_var.set(
            f"{star}Selected preset '{preset.get('name', '')}' -> {preset.get('mode', '')} using {preset.get('engine_mode', ENGINE_AUTO)}."
        )

    def _refresh_presets_view(self) -> None:
        if not hasattr(self, "preset_tree"):
            return
        self.preset_item_ids.clear()
        for item in self.preset_tree.get_children():
            self.preset_tree.delete(item)
        presets = self.state_store.presets()
        presets.sort(key=lambda item: (not bool(item.get("favorite", False)), str(item.get("name", "")).lower()))
        for index, preset in enumerate(presets):
            iid = f"preset_{index}"
            values = (
                "★" if preset.get("favorite") else "",
                preset.get("name", ""),
                preset.get("mode", ""),
                preset.get("engine_mode", ""),
                "Yes" if preset.get("merge_to_one_pdf") else "No",
            )
            self.preset_tree.insert("", tk.END, iid=iid, values=values)
            self.preset_item_ids[iid] = preset
        self._refresh_automation_page()
        self._refresh_favorite_preset_widgets()

    def _current_preset_record(self) -> dict[str, object]:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        existing = None
        for item in self.state_store.presets():
            if str(item.get("name", "")).strip().lower() == self.preset_name_var.get().strip().lower():
                existing = item
                break
        return {
            "name": self.preset_name_var.get().strip() or f"{self.mode_var.get()} preset",
            "mode": self.mode_var.get(),
            "output_dir": self.output_dir_var.get().strip(),
            "merge_to_one_pdf": bool(self.merge_var.get()),
            "merged_output_name": self.output_name_var.get().strip() or default_merged_name(self.mode_var.get()),
            "image_format": self.image_format_var.get().strip() or "png",
            "image_scale": self.image_scale_var.get().strip() or "2.0",
            "engine_mode": self.engine_mode_var.get().strip().lower() or ENGINE_AUTO,
            "recursive": bool(self.recursive_var.get()),
            "favorite": bool((existing or {}).get("favorite", False)),
            "created_at": str((existing or {}).get("created_at", now)),
            "updated_at": now,
        }

    def _save_current_preset(self) -> None:
        name = self.preset_name_var.get().strip()
        if not name:
            messagebox.showwarning("Preset name", "Enter a preset name first.")
            return
        self.state_store.save_preset(self._current_preset_record())
        self._refresh_presets_view()
        self._log_automation(f"Saved preset '{name}'.")

    def _apply_selected_preset(self) -> None:
        preset = self._selected_preset()
        if not preset:
            messagebox.showinfo("Presets", "Select a preset first.")
            return
        self._apply_preset_record(preset, start=False)

    def _delete_selected_preset(self) -> None:
        preset = self._selected_preset()
        if not preset:
            messagebox.showinfo("Presets", "Select a preset first.")
            return
        name = str(preset.get("name", "")).strip()
        if not messagebox.askyesno("Delete preset", f"Remove preset '{name}' from this machine?"):
            return
        self.state_store.delete_preset(name)
        self._refresh_presets_view()
        self._log_automation(f"Deleted preset '{name}'.")

    def _export_presets(self) -> None:
        presets = self.state_store.presets()
        if not presets:
            messagebox.showinfo("Presets", "There are no presets to export yet.")
            return
        file_path = filedialog.asksaveasfilename(
            title="Export presets to JSON",
            defaultextension=".json",
            initialfile="gokul_omni_presets.json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not file_path:
            return
        export_presets_to_json(presets, Path(file_path))
        self._log_automation(f"Exported {len(presets)} preset(s) to {Path(file_path).name}.")

    def _import_presets(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Import presets from JSON",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not file_path:
            return
        imported = import_presets_from_json(Path(file_path))
        if not imported:
            messagebox.showwarning("Import presets", "No preset records were found in that JSON file.")
            return
        merged: dict[str, dict[str, object]] = {
            str(item.get("name", "")).strip().lower(): item
            for item in self.state_store.presets()
            if str(item.get("name", "")).strip()
        }
        for item in imported:
            key = str(item.get("name", "")).strip().lower()
            if key:
                merged[key] = item
        self.state_store.replace_presets(list(merged.values()))
        self._refresh_presets_view()
        self._log_automation(f"Imported {len(imported)} preset(s) from {Path(file_path).name}.")

    def _current_watch_config(self) -> dict[str, object]:
        output_dir = self.watch_output_dir_var.get().strip() or self.output_dir_var.get().strip() or str(Path.cwd() / "converted_output")
        payload = WatchFolderConfig(
            source_dir=self.watch_source_dir_var.get().strip(),
            output_dir=output_dir,
            mode=self.watch_mode_var.get().strip() or MODE_ANY_TO_PDF,
            merge_to_one_pdf=bool(self.watch_merge_var.get()),
            merged_output_name=self.watch_output_name_var.get().strip() or "watch_output",
            recursive=bool(self.watch_recursive_var.get()),
            interval_seconds=int(self.watch_interval_var.get().strip() or "15") if self.watch_interval_var.get().strip().isdigit() else 15,
            engine_mode=self.watch_engine_var.get().strip().lower() or ENGINE_AUTO,
            archive_processed=bool(self.watch_archive_var.get()),
            archive_dir=self.watch_archive_dir_var.get().strip(),
            create_zip_bundle=bool(self.watch_zip_var.get()),
            create_report=bool(self.watch_report_var.get()),
            open_mail_draft=bool(self.watch_mail_var.get()),
            skip_existing_on_start=bool(self.watch_skip_existing_var.get()),
        )
        return normalize_watch_config(payload.to_dict()).to_dict()

    def _save_watch_config(self) -> None:
        config = self._current_watch_config()
        if not str(config.get("source_dir", "")).strip():
            messagebox.showwarning("Watch folder", "Choose a source watch folder first.")
            return
        self.state_store.set_watch_config(config)
        self.watch_status_var.set("Watch-folder settings saved locally.")
        self._refresh_automation_page()
        self._log_automation("Saved watch-folder automation settings.")

    def _browse_watch_source_dir(self) -> None:
        folder = filedialog.askdirectory(title="Choose watch source folder")
        if folder:
            self.watch_source_dir_var.set(folder)
            self._refresh_automation_page()

    def _browse_watch_output_dir(self) -> None:
        folder = filedialog.askdirectory(title="Choose automation output folder")
        if folder:
            self.watch_output_dir_var.set(folder)
            self._refresh_automation_page()

    def _browse_watch_archive_dir(self) -> None:
        folder = filedialog.askdirectory(title="Choose archive folder for processed source files")
        if folder:
            self.watch_archive_dir_var.set(folder)

    def _clear_watch_seen_cache(self) -> None:
        if not self.state_store.watch_seen_fingerprints():
            return
        if not messagebox.askyesno("Reset seen cache", "Clear the watch-folder memory of already processed files?"):
            return
        self.state_store.clear_watch_seen()
        self._refresh_automation_page()
        self._log_automation("Cleared the watch-folder seen-file cache.")

    def _schedule_automation_tick(self, delay_ms: int | None = None) -> None:
        if not self.automation_active:
            return
        if self.automation_after_id:
            try:
                self.after_cancel(self.automation_after_id)
            except Exception:
                pass
            self.automation_after_id = None
        config = normalize_watch_config(self._current_watch_config())
        interval = delay_ms if delay_ms is not None else max(2000, int(config.interval_seconds) * 1000)
        self.automation_after_id = self.after(interval, self._automation_tick)

    def _start_watch_automation(self) -> None:
        config = normalize_watch_config(self._current_watch_config())
        if not config.source_dir:
            messagebox.showwarning("Watch folder", "Choose a source watch folder first.")
            return
        source_dir = Path(config.source_dir).expanduser()
        if not source_dir.exists() or not source_dir.is_dir():
            messagebox.showerror("Watch folder", "The selected watch folder does not exist.")
            return
        self.state_store.set_watch_config(config.to_dict())
        self.automation_active = True
        self.watch_status_var.set(f"Watching {source_dir} every {config.interval_seconds} second(s).")
        if config.skip_existing_on_start:
            existing_files, existing_fingerprints = discover_watch_candidates(
                source_dir,
                config.mode,
                config.recursive,
                self.state_store.watch_seen_fingerprints(),
            )
            if existing_fingerprints:
                self.state_store.add_watch_seen(existing_fingerprints)
                self._log_automation(
                    f"Watcher started and cached {len(existing_files)} existing file(s) so only new arrivals will be processed."
                )
            else:
                self._log_automation("Watcher started. No existing files needed to be cached.")
        else:
            self._log_automation("Watcher started. Existing matching files can be picked up on the next scan.")
        self._refresh_automation_page()
        self._schedule_automation_tick(1000)

    def _stop_watch_automation(self) -> None:
        self.automation_active = False
        if self.automation_after_id:
            try:
                self.after_cancel(self.automation_after_id)
            except Exception:
                pass
            self.automation_after_id = None
        self.watch_status_var.set("Automation watcher stopped.")
        self._refresh_automation_page()
        self._log_automation("Watcher stopped.")

    def _automation_tick(self) -> None:
        self.automation_after_id = None
        if not self.automation_active:
            return
        if self.running:
            self.watch_status_var.set("Watcher is waiting for the current job to finish.")
            self._schedule_automation_tick()
            return
        config = normalize_watch_config(self._current_watch_config())
        source_dir = Path(config.source_dir).expanduser()
        if not source_dir.exists() or not source_dir.is_dir():
            self.watch_status_var.set("Watch folder is missing. Automation was paused.")
            self.automation_active = False
            self._refresh_automation_page()
            return
        files, fingerprints = discover_watch_candidates(
            source_dir,
            config.mode,
            config.recursive,
            self.state_store.watch_seen_fingerprints(),
        )
        if not files:
            self.watch_status_var.set(f"Watching {source_dir.name}. No new files found yet.")
            self._schedule_automation_tick()
            return
        self._log_automation(f"Detected {len(files)} new file(s) in the watch folder. Starting automation run.")
        self._start_automation_batch(files, fingerprints, config)

    def _scan_watch_folder_now(self) -> None:
        if self.running:
            messagebox.showinfo("Automation", "A job is already running. Scan again after it finishes.")
            return
        config = normalize_watch_config(self._current_watch_config())
        if not config.source_dir:
            messagebox.showwarning("Watch folder", "Choose a source watch folder first.")
            return
        source_dir = Path(config.source_dir).expanduser()
        files, fingerprints = discover_watch_candidates(
            source_dir,
            config.mode,
            config.recursive,
            self.state_store.watch_seen_fingerprints(),
        )
        if not files:
            self.watch_status_var.set("No new matching files were found in the watch folder.")
            self._refresh_automation_page()
            return
        self._log_automation(f"Manual scan found {len(files)} file(s). Starting automation run.")
        self._start_automation_batch(files, fingerprints, config)

    def _run_batch_config(self, config: BatchConfig, origin: str, heading: str, target_page: str) -> None:
        self.running = True
        self.active_run_kind = origin
        self.last_run_origin = origin
        self._set_ocr_buttons_enabled(False)
        self.progress["value"] = 0
        self.status_var.set("Working...")
        self._append_log(heading)
        self._append_log(f"Mode: {config.mode}")
        self._append_log(f"Engine: {config.engine_mode}")
        self._append_log(
            build_conversion_route_preview(
                config.mode,
                config.files,
                engine_mode=config.engine_mode,
                soffice_path=config.soffice_path,
            )
        )
        if config.soffice_path:
            self._append_log(f"Configured soffice path: {config.soffice_path}")
        self._append_log(f"Output folder: {config.output_dir}")
        self._append_log(f"Files selected: {len(config.files)}")
        worker = threading.Thread(target=self._worker_run, args=(config,), daemon=True)
        worker.start()
        self.after(100, self._poll_worker_queue)
        self._show_page(target_page)
        self._refresh_home_summary()

    def _start_conversion(self) -> None:
        if self.running:
            return
        if not self.selected_files:
            messagebox.showwarning("No files selected", "Please add some input files or a folder first.")
            self._show_page("convert")
            return

        try:
            image_scale = float(self.image_scale_var.get().strip() or "2.0")
            image_scale = self._effective_image_scale(image_scale)
        except ValueError:
            messagebox.showerror("Invalid image scale", "Image scale must be a number such as 1.5 or 2.0.")
            self._show_page("convert")
            return

        output_dir_text = self.output_dir_var.get().strip()
        if not output_dir_text:
            output_dir = self.selected_files[0].parent / "converted_output"
            self.output_dir_var.set(str(output_dir))
        else:
            output_dir = Path(output_dir_text).expanduser()

        config = BatchConfig(
            mode=self.mode_var.get(),
            files=self.selected_files.copy(),
            output_dir=output_dir,
            merge_to_one_pdf=self.merge_var.get(),
            merged_output_name=self.output_name_var.get().strip() or default_merged_name(self.mode_var.get()),
            image_format=self.image_format_var.get().strip() or "png",
            image_scale=image_scale,
            engine_mode=self.engine_mode_var.get().strip().lower() or ENGINE_AUTO,
            soffice_path=self.soffice_path_var.get().strip(),
        )
        self._run_batch_config(config, origin="convert", heading="\n=== New run started ===", target_page="convert")

    def _start_automation_batch(
        self,
        files: list[Path],
        fingerprints: list[str],
        watch_config: WatchFolderConfig,
    ) -> None:
        if self.running:
            return
        output_dir = Path(watch_config.output_dir).expanduser()
        self.automation_current_files = files.copy()
        self.automation_current_fingerprints = fingerprints.copy()
        config = BatchConfig(
            mode=watch_config.mode,
            files=files.copy(),
            output_dir=output_dir,
            merge_to_one_pdf=watch_config.merge_to_one_pdf,
            merged_output_name=watch_config.merged_output_name,
            image_format=self.image_format_var.get().strip() or "png",
            image_scale=float(self.image_scale_var.get().strip() or "2.0"),
            engine_mode=watch_config.engine_mode,
            soffice_path=self.soffice_path_var.get().strip(),
        )
        self.watch_status_var.set(f"Automation is processing {len(files)} file(s) from the watch folder.")
        self._run_batch_config(config, origin="automation", heading="\n=== Automation run started ===", target_page="automation")

    def _open_last_output_folder(self) -> None:
        if self.last_output_dir is not None:
            open_path(self.last_output_dir)
            return
        self._open_output_folder()

    def _create_zip_bundle_for_last_outputs(self) -> None:
        if not self.last_outputs:
            messagebox.showinfo("ZIP bundle", "Run a conversion or PDF tool first so the app has outputs to bundle.")
            return
        output_dir = self.last_output_dir or Path(self.output_dir_var.get().strip() or Path.cwd() / "converted_output")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = self.share_bundle_name_var.get().strip() or "gokul_outputs_bundle"
        destination = output_dir / f"{base}_{timestamp}.zip"
        bundle_paths_as_zip(self.last_outputs, destination)
        self.last_outputs = [*self.last_outputs, destination]
        self.last_output_dir = output_dir
        self.status_var.set(f"Created ZIP bundle: {destination.name}")
        self._log_automation(f"Created ZIP bundle for last outputs: {destination.name}")

    def _export_last_run_report(self) -> None:
        record = self.last_job_record or (self.state_store.recent_jobs()[0] if self.state_store.recent_jobs() else None)
        if not record:
            messagebox.showinfo("Run report", "Run a conversion or PDF tool first so the app has a job summary to export.")
            return
        output_dir = self.last_output_dir or Path(str(record.get("output_dir", self.output_dir_var.get()))).expanduser()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = self.share_report_name_var.get().strip() or "gokul_last_run_report"
        destination = output_dir / f"{base}_{timestamp}.txt"
        write_run_report(dict(record), destination)
        if destination not in self.last_outputs:
            self.last_outputs = [*self.last_outputs, destination]
        self.last_output_dir = output_dir
        self.status_var.set(f"Saved run report: {destination.name}")
        self._log_automation(f"Saved last-run report: {destination.name}")

    def _record_external_job(self, record: dict[str, object]) -> None:
        self.last_job_record = dict(record)
        self.state_store.add_recent_job(record)
        self._refresh_history_views()

    def _record_job(
        self,
        config: BatchConfig | PdfToolConfig,
        outputs: list[Path],
        status: str,
        error_text: str = "",
    ) -> None:
        preview = [str(path) for path in getattr(config, "files", [])[:8]]
        if isinstance(config, PdfToolConfig):
            record = {
                "job_type": "pdf_tool",
                "status": status,
                "mode": f"PDF Tool -> {config.tool}",
                "tool": config.tool,
                "file_count": len(config.files),
                "output_count": len(outputs),
                "output_dir": str(config.output_dir),
                "output_name": config.output_name,
                "page_spec": config.page_spec,
                "every_n_pages": config.every_n_pages,
                "watermark_text": config.watermark_text,
                "watermark_image": str(config.watermark_image) if config.watermark_image else "",
                "watermark_font_size": config.watermark_font_size,
                "watermark_rotation": config.watermark_rotation,
                "watermark_opacity": config.watermark_opacity,
                "watermark_position": config.watermark_position,
                "watermark_image_scale_percent": config.watermark_image_scale_percent,
                "metadata_title": config.metadata_title,
                "metadata_author": config.metadata_author,
                "metadata_subject": config.metadata_subject,
                "metadata_keywords": config.metadata_keywords,
                "metadata_clear_existing": config.metadata_clear_existing,
                "compression_profile": config.compression_profile,
                "password_used": bool(config.pdf_password),
                "owner_password_used": bool(config.pdf_owner_password),
                "redact_rect": config.redact_rect,
                "replacement_text": config.replacement_text,
                "input_files": [str(path) for path in config.files],
                "inputs_preview": preview,
                "outputs_preview": [str(path) for path in outputs[:8]],
                "error": error_text or "",
            }
        else:
            record = {
                "job_type": "convert",
                "status": status,
                "mode": config.mode,
                "file_count": len(config.files),
                "output_count": len(outputs),
                "output_dir": str(config.output_dir),
                "merge_to_one_pdf": config.merge_to_one_pdf,
                "merged_output_name": config.merged_output_name,
                "image_format": config.image_format,
                "image_scale": config.image_scale,
                "engine_mode": config.engine_mode,
                "soffice_configured": bool(config.soffice_path),
                "run_origin": self.last_run_origin or "manual",
                "inputs_preview": preview,
                "outputs_preview": [str(path) for path in outputs[:8]],
                "error": error_text or "",
            }
        self.last_job_record = dict(record)
        self.state_store.add_recent_job(record)
        self._refresh_history_views()

    def _handle_automation_run_completion(
        self,
        success: bool,
        config: BatchConfig,
        outputs: list[Path] | None = None,
        error_text: str = "",
    ) -> None:
        current_files = self.automation_current_files.copy()
        current_fingerprints = self.automation_current_fingerprints.copy()
        self.automation_current_files = []
        self.automation_current_fingerprints = []
        if current_fingerprints:
            self.state_store.add_watch_seen(current_fingerprints)

        watch_config = normalize_watch_config(self._current_watch_config())
        output_paths = [Path(item) for item in (outputs or [])]
        if success:
            extra_outputs: list[Path] = []
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            if watch_config.create_report and self.last_job_record:
                report_base = self.share_report_name_var.get().strip() or "watch_run_report"
                report_path = write_run_report(self.last_job_record, Path(config.output_dir) / f"{report_base}_{timestamp}.txt")
                extra_outputs.append(report_path)
                self._log_automation(f"Saved automation report: {report_path.name}")
            if watch_config.create_zip_bundle and (output_paths or extra_outputs):
                bundle_base = self.share_bundle_name_var.get().strip() or "watch_bundle"
                bundle_path = bundle_paths_as_zip([*output_paths, *extra_outputs], Path(config.output_dir) / f"{bundle_base}_{timestamp}.zip")
                extra_outputs.append(bundle_path)
                self._log_automation(f"Created automation ZIP bundle: {bundle_path.name}")
            if watch_config.archive_processed:
                source_root = Path(watch_config.source_dir).expanduser()
                archive_root = Path(watch_config.archive_dir).expanduser() if watch_config.archive_dir.strip() else source_root / "processed"
                moved = move_files_to_archive(current_files, source_root, archive_root)
                self._log_automation(f"Archived {len(moved)} processed source file(s) to {archive_root}.")
            combined_outputs = [*output_paths, *extra_outputs]
            if combined_outputs:
                self.last_outputs = combined_outputs
                self.last_output_dir = Path(config.output_dir)
                self.last_job_label = f"Automation -> {config.mode}"
            self.watch_status_var.set(f"Automation finished for {len(current_files)} watched file(s).")
            self.state_store.add_automation_event({
                "status": "Completed",
                "mode": config.mode,
                "message": f"Processed {len(current_files)} file(s) from the watch folder using {config.mode}.",
            })
            if watch_config.open_mail_draft and combined_outputs:
                self.after(200, self._open_mail_draft_for_last_outputs)
        else:
            self.watch_status_var.set(
                "Automation failed for the latest watch-folder batch. Seen-cache entries were still stored to avoid endless retry loops."
            )
            self.state_store.add_automation_event({
                "status": "Failed",
                "mode": config.mode,
                "message": f"Automation failed for {len(current_files)} file(s). Reset the seen cache to retry them. {error_text}".strip(),
            })
            self._log_automation(
                f"Automation failed for {len(current_files)} file(s). Reset the seen cache if you want to retry them.",
                persist=False,
            )
        self._refresh_automation_page()
        if self.automation_active:
            self._schedule_automation_tick()

    def _poll_worker_queue(self) -> None:
        keep_polling = self.running
        try:
            while True:
                message_type, payload = self.worker_queue.get_nowait()
                if message_type == "log":
                    self._append_log(str(payload))
                elif message_type == "progress":
                    current, total = payload  # type: ignore[misc]
                    total = max(int(total), 1)
                    current = int(current)
                    self.progress["value"] = round((current / total) * 100, 2)
                    self.status_var.set(f"Processing {current}/{total}...")
                elif message_type == "done":
                    config, outputs = payload  # type: ignore[assignment]
                    run_kind = self.active_run_kind
                    self.running = False
                    self.active_run_kind = ""
                    keep_polling = False
                    self._set_ocr_buttons_enabled(True)
                    self.progress["value"] = 100
                    self.status_var.set("Completed.")
                    self.last_outputs = [Path(item) for item in outputs]
                    self.last_output_dir = Path(str(config.output_dir))
                    self.last_job_label = str(config.mode)
                    self._append_log("Completed successfully.")
                    for item in outputs:
                        self._append_log(f"Created: {item}")
                    self._record_job(config, outputs=outputs, status="Completed")
                    self._refresh_home_summary()
                    if run_kind == "automation":
                        self._handle_automation_run_completion(True, config, outputs=outputs)
                    else:
                        messagebox.showinfo(
                            "Done",
                            f"Conversion finished successfully.\n\nOutput folder:\n{self.output_dir_var.get()}",
                        )
                elif message_type == "error":
                    config, details = payload  # type: ignore[assignment]
                    run_kind = self.active_run_kind
                    self.running = False
                    self.active_run_kind = ""
                    keep_polling = False
                    self._set_ocr_buttons_enabled(True)
                    self.status_var.set("Failed.")
                    self._append_log("ERROR:\n" + str(details))
                    self._record_job(config, outputs=[], status="Failed", error_text=str(details))
                    self._refresh_home_summary()
                    if run_kind == "automation":
                        self._handle_automation_run_completion(False, config, outputs=[], error_text=str(details))
                    else:
                        messagebox.showerror("Conversion failed", str(details))
                elif message_type == "pdf_log":
                    self._append_pdf_tool_log(str(payload))
                elif message_type == "pdf_progress":
                    current, total = payload  # type: ignore[misc]
                    total = max(int(total), 1)
                    current = int(current)
                    self.pdf_tool_progress["value"] = round((current / total) * 100, 2)
                    self.status_var.set(f"Running PDF tool {current}/{total}...")
                elif message_type == "pdf_done":
                    config, outputs = payload  # type: ignore[assignment]
                    self.running = False
                    self.active_run_kind = ""
                    keep_polling = False
                    self._set_ocr_buttons_enabled(True)
                    self.pdf_tool_progress["value"] = 100
                    self.status_var.set("PDF tool completed.")
                    self._track_success_outputs(outputs, output_dir=Path(str(config.output_dir)), label=f"PDF Tool -> {config.tool}")
                    self._append_pdf_tool_log("Completed successfully.")
                    self._record_job(config, outputs=outputs, status="Completed")
                    self._maybe_auto_open_output(Path(str(config.output_dir)))
                    messagebox.showinfo(
                        "Done",
                        f"PDF tool finished successfully.\n\nOutput folder:\n{self.output_dir_var.get()}",
                    )
                elif message_type == "pdf_error":
                    config, details = payload  # type: ignore[assignment]
                    self.running = False
                    self.active_run_kind = ""
                    keep_polling = False
                    self._set_ocr_buttons_enabled(True)
                    self.status_var.set("PDF tool failed.")
                    self._append_pdf_tool_log("ERROR:\n" + str(details))
                    self._record_job(config, outputs=[], status="Failed", error_text=str(details))
                    messagebox.showerror("PDF tool failed", str(details))
                elif message_type == "ocr_log":
                    self._append_ocr_log(str(payload))
                elif message_type == "ocr_progress":
                    current, total = payload  # type: ignore[misc]
                    total = max(int(total), 1)
                    current = int(current)
                    self.ocr_progress["value"] = round((current / total) * 100, 2)
                    self.status_var.set(f"Running OCR {current}/{total}...")
                elif message_type == "ocr_done":
                    record, outputs = payload  # type: ignore[assignment]
                    self.running = False
                    self.active_run_kind = ""
                    keep_polling = False
                    self._set_ocr_buttons_enabled(True)
                    self.ocr_progress["value"] = 100
                    self.status_var.set("OCR completed.")
                    self.ocr_status_var.set("OCR task completed successfully.")
                    output_dir = Path(str(record.get("output_dir", self.ocr_output_var.get())))
                    self._track_success_outputs(outputs, output_dir=output_dir, label=str(record.get("mode", "OCR")))
                    self._append_ocr_log("Completed successfully.")
                    for item in outputs:
                        self._append_ocr_log(f"Created: {item}")
                    completed_record = dict(record)
                    completed_record["status"] = "Completed"
                    completed_record["output_count"] = len(outputs)
                    completed_record["outputs_preview"] = [str(item) for item in outputs[:20]]
                    self._record_external_job(completed_record)
                    self._refresh_home_summary()
                    self._maybe_auto_open_output(output_dir)
                    messagebox.showinfo(
                        "OCR complete",
                        f"OCR finished successfully.\n\nOutput folder:\n{output_dir}",
                    )
                elif message_type == "ocr_error":
                    record, details = payload  # type: ignore[assignment]
                    self.running = False
                    self.active_run_kind = ""
                    keep_polling = False
                    self._set_ocr_buttons_enabled(True)
                    self.status_var.set("OCR failed.")
                    self.ocr_status_var.set("OCR failed. Review the log for details.")
                    self._append_ocr_log("ERROR:\n" + str(details))
                    failed_record = dict(record)
                    failed_record["status"] = "Failed"
                    failed_record["error"] = str(details)
                    failed_record["output_count"] = 0
                    failed_record["outputs_preview"] = []
                    self._record_external_job(failed_record)
                    self._refresh_home_summary()
                    self._refresh_failed_jobs_view()
                    messagebox.showerror("OCR failed", str(details))
        except queue.Empty:
            pass

        if keep_polling:
            self.after(100, self._poll_worker_queue)

    def _persist_state(self) -> None:
        try:
            smtp_settings = self._smtp_settings_from_vars().to_state_dict()
        except Exception:
            smtp_settings = SMTPSettings.from_dict(self.state_store.get("smtp_settings", {})).to_state_dict()
        try:
            link_timeout = max(int(str(self.link_timeout_var.get()).strip() or "25"), 5)
        except Exception:
            link_timeout = 25
        ensure_install_date(self.state_store.state)
        session_snapshot = self._capture_session_snapshot() if bool(self.restore_session_var.get()) else {}
        try:
            cache_keep_days = max(0, int(str(self.link_cache_max_age_var.get()).strip() or "0"))
        except Exception:
            cache_keep_days = 0
        try:
            cache_size_mb = max(32, int(str(self.link_cache_max_size_var.get()).strip() or "512"))
        except Exception:
            cache_size_mb = 512
        self.state_store.update(
            theme=self.theme_choice_var.get().strip().lower() or "dark",
            output_dir=self.output_dir_var.get().strip() or str(Path.cwd() / "converted_output"),
            recursive_scan=bool(self.recursive_var.get()),
            conversion_engine=self.engine_mode_var.get().strip().lower() or ENGINE_AUTO,
            performance_mode=self.performance_mode_var.get().strip().lower() or "balanced",
            soffice_path=self.soffice_path_var.get().strip(),
            tesseract_path=self.tesseract_path_var.get().strip(),
            ocr_language=self.ocr_language_var.get().strip() or "eng",
            ocr_dpi=int(str(self.ocr_dpi_var.get()).strip() or "220"),
            ocr_psm=int(str(self.ocr_psm_var.get()).strip() or "6"),
            ocr_output_dir=self.ocr_output_var.get().strip(),
            smtp_settings=smtp_settings,
            install_date=self.state_store.get("install_date", ""),
            splash_enabled=bool(self.splash_enabled_var.get()),
            splash_gif_path=self.splash_gif_path_var.get().strip() or "assets/gokul_splash.gif",
            login_popup_enabled=bool(self.login_popup_enabled_var.get()),
            login_popup_dismissed=bool(self.state_store.get("login_popup_dismissed", False)),
            login_popup_completed=bool(self.state_store.get("login_popup_completed", False)),
            login_popup_last_shown=str(self.state_store.get("login_popup_last_shown", "")),
            splash_seen=bool(self.state_store.get("splash_seen", False)),
            link_cache_dir=self.link_cache_dir_var.get().strip(),
            link_timeout=link_timeout,
            link_keep_downloads=bool(self.link_keep_downloads_var.get()),
            link_cache_max_age_days=cache_keep_days,
            link_cache_max_size_mb=cache_size_mb,
            recent_links=self.state_store.recent_links(),
            auto_open_output_folder=bool(self.auto_open_output_var.get()),
            restore_last_session=bool(self.restore_session_var.get()),
            cleanup_temp_on_exit=bool(self.cleanup_temp_var.get()),
            update_checker_enabled=bool(self.update_checker_enabled_var.get()),
            last_update_check=self.last_update_check_var.get().strip(),
            recent_outputs=self.state_store.recent_outputs(),
            failed_jobs=self.state_store.failed_jobs(),
            session_snapshot=session_snapshot,
            window_geometry=self.geometry(),
            last_page=self.current_page,
        )
        self.state_store.set_watch_config(self._current_watch_config())

    def _on_close(self) -> None:
        if self.running and not messagebox.askyesno(
            "Exit",
            "A conversion is still running. Close the app anyway?",
        ):
            return
        self.automation_active = False
        if self.automation_after_id:
            try:
                self.after_cancel(self.automation_after_id)
            except Exception:
                pass
            self.automation_after_id = None
        self._persist_state()
        for window in (
            self.notes_window,
            self.smtp_window,
            self.build_center_window,
            self.about_editor_window,
            self.command_palette_window,
            self.splash_window,
            self.login_popup_window,
        ):
            if window and window.winfo_exists():
                window.destroy()
        if bool(self.cleanup_temp_var.get()):
            shutil.rmtree(self.session_temp_root, ignore_errors=True)
        self.destroy()


def main() -> None:
    parser = argparse.ArgumentParser(description="Gokul Omni Convert Lite")
    parser.add_argument("--skip-startup-overlays", action="store_true", help="Skip splash and reminder overlays on startup.")
    args = parser.parse_args()
    app = GokulOmniConvertLiteApp(skip_startup_overlays=args.skip_startup_overlays)
    app.mainloop()


if __name__ == "__main__":
    main()
