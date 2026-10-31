"""Settings page: appearance, engine, local files and folders, and quick actions."""

from __future__ import annotations

import tkinter as tk


def _readonly(kit, parent, value: str):
    """An entry that shows a path but cannot be edited."""
    variable = tk.StringVar(value=str(value))
    entry = kit.entry(parent, variable)
    entry.configure(state="disabled")
    entry._readonly_var = variable  # keep a reference so Tk does not collect it
    return entry


def build(app) -> None:
    kit = app.kit
    page = app._create_page_frame("settings", scrollable=True)
    page.grid_columnconfigure(0, weight=1)
    page.grid_rowconfigure(1, weight=1)

    kit.page_header(
        page,
        "Settings",
        "Appearance, conversion engine, local files, and optional remote branding assets. "
        "Everything is saved automatically when you close the app.",
    ).grid(row=0, column=0, sticky="ew", pady=(0, kit.pad("lg")))

    content = kit.frame(page, variant="transparent")
    app.settings_content_frame = content
    content.grid(row=1, column=0, sticky="nsew")
    content.grid_columnconfigure(0, weight=1)
    content.grid_columnconfigure(1, weight=1)
    content.grid_rowconfigure(0, weight=1)
    content.grid_rowconfigure(1, weight=1)

    _build_appearance(app, kit, content)
    _build_engine(app, kit, content)
    _build_files(app, kit, content)
    _build_quick_actions(app, kit, content)


def _build_appearance(app, kit, content) -> None:
    from app_state import UI_SCALE_OPTIONS

    card = kit.card(content, title="Appearance & behaviour")
    app.settings_appearance_frame = card
    card.grid(row=0, column=0, sticky="nsew", padx=(0, kit.pad("md")), pady=(0, kit.pad("md")))
    body = card.body
    body.grid_columnconfigure(1, weight=1)

    kit.label(body, "Theme", variant="body").grid(row=0, column=0, sticky="w", padx=(0, kit.pad("sm")))
    kit.combo(body, ["dark", "light", "system"], app.theme_choice_var).grid(row=0, column=1, sticky="w")
    kit.button(body, "Apply", app._apply_theme, variant="secondary", width=90).grid(
        row=0, column=2, padx=(kit.pad("sm"), 0)
    )

    row = 1
    kit.label(body, "UI scale", variant="body").grid(
        row=row, column=0, sticky="w", padx=(0, kit.pad("sm")), pady=(kit.pad("sm"), 0)
    )
    app.ui_scale_combo = kit.combo(
        body,
        sorted(UI_SCALE_OPTIONS, key=lambda item: int(item.rstrip("%"))),
        app.ui_scale_var,
    )
    app.ui_scale_combo.grid(row=row, column=1, sticky="w", pady=(kit.pad("sm"), 0))
    app.ui_scale_combo.bind("<<ComboboxSelected>>", lambda _event: app._apply_theme())
    row += 1

    kit.label(body, "Start page", variant="body").grid(
        row=row, column=0, sticky="w", padx=(0, kit.pad("sm")), pady=(kit.pad("sm"), 0)
    )
    kit.combo(
        body,
        [
            "home", "convert", "pdf_tools", "ocr", "organizer",
            "archive", "rename_organize", "automation", "history", "settings", "about",
        ],
        app.start_page_var,
    ).grid(row=row, column=1, sticky="w", pady=(kit.pad("sm"), 0))
    row += 1

    kit.labelled(
        body, "Performance",
        "Eco caps heavy image rendering, which helps on older machines.\n"
        "Balanced is the default.\nQuality keeps your requested scale untouched.",
    ).grid(row=row, column=0, sticky="w", padx=(0, kit.pad("sm")), pady=(kit.pad("sm"), 0))
    kit.combo(body, ["eco", "balanced", "quality"], app.performance_mode_var).grid(
        row=row, column=1, sticky="w", pady=(kit.pad("sm"), 0)
    )
    row += 1

    kit.separator(body).grid(row=row, column=0, columnspan=3, sticky="ew", pady=kit.pad("md"))
    row += 1

    for text, variable, command in (
        ("Use compact UI spacing", app.compact_ui_var, app._apply_theme),
        ("High contrast", app.high_contrast_var, app._apply_theme),
        ("Reduced motion", app.reduced_motion_var, None),
        ("Scan folders recursively by default", app.recursive_var, None),
        ("Show the first-launch splash", app.splash_enabled_var, None),
        ("Show the login reminder", app.login_popup_enabled_var, None),
        ("Open the output folder after a successful run", app.auto_open_output_var, None),
        ("Restore the last session on startup", app.restore_session_var, None),
        ("Clean temporary session files on exit", app.cleanup_temp_var, None),
        ("Enable update-checker reminders", app.update_checker_enabled_var, None),
        ("Back up state automatically before saving", app.state_backup_enabled_var, None),
    ):
        kwargs = {"command": command} if command else {}
        kit.check(body, text, variable, **kwargs).grid(
            row=row, column=0, columnspan=3, sticky="w", pady=(kit.pad("xs"), 0)
        )
        row += 1

    kit.label(body, "Keep backups", variant="body").grid(
        row=row, column=0, sticky="w", padx=(0, kit.pad("sm")), pady=(kit.pad("md"), 0)
    )
    kit.entry(body, app.state_backup_keep_var, width=90).grid(row=row, column=1, sticky="w", pady=(kit.pad("md"), 0))
    row += 1

    for variable in (app.backup_summary_var, app.login_popup_state_var):
        kit.label(
            body, "", variant="small", textvariable=variable, wraplength=420, justify="left", anchor="w"
        ).grid(row=row, column=0, columnspan=3, sticky="ew", pady=(kit.pad("sm"), 0))
        row += 1

    backup_actions = kit.button_bar(body, columns=2)
    backup_actions.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(kit.pad("md"), 0))
    backup_actions.button("Create backup now", app._backup_state_now)
    backup_actions.button("Restore latest", app._restore_latest_state_backup_action)
    backup_actions.button("Open backup folder", app._open_state_backup_dir, variant="ghost")


def _build_engine(app, kit, content) -> None:
    from converter_core import ENGINE_ORDER

    card = kit.card(content, title="Conversion engine")
    app.settings_engine_frame = card
    card.grid(row=0, column=1, sticky="nsew", pady=(0, kit.pad("md")))
    body = card.body
    body.grid_columnconfigure(1, weight=1)

    kit.label(body, "Engine", variant="body").grid(row=0, column=0, sticky="w", padx=(0, kit.pad("sm")))
    app.engine_combo = kit.combo(body, list(ENGINE_ORDER), app.engine_mode_var)
    app.engine_combo.grid(row=0, column=1, sticky="w")
    app.engine_combo.bind("<<ComboboxSelected>>", app._on_engine_changed)
    kit.button(body, "Refresh status", app._refresh_dependency_status, variant="ghost").grid(
        row=0, column=2, padx=(kit.pad("sm"), 0)
    )

    kit.label(
        body, "", variant="small", textvariable=app.engine_help_var, wraplength=420, justify="left", anchor="w"
    ).grid(row=1, column=0, columnspan=3, sticky="ew", pady=(kit.pad("sm"), 0))

    kit.label(body, "LibreOffice", variant="body").grid(
        row=2, column=0, sticky="w", padx=(0, kit.pad("sm")), pady=(kit.pad("md"), 0)
    )
    app.soffice_entry = kit.entry(body, app.soffice_path_var, placeholder="Auto-detected if installed")
    app.soffice_entry.grid(row=2, column=1, columnspan=2, sticky="ew", pady=(kit.pad("md"), 0))
    soffice_actions = kit.button_bar(body, columns=3)
    soffice_actions.grid(row=3, column=1, columnspan=2, sticky="w", pady=(kit.pad("sm"), 0))
    soffice_actions.button("Browse", app._browse_soffice_path)
    soffice_actions.button("Test", app._test_soffice_path)
    soffice_actions.button("Clear", app._clear_soffice_path, variant="ghost")

    kit.label(body, "Tesseract", variant="body").grid(
        row=4, column=0, sticky="w", padx=(0, kit.pad("sm")), pady=(kit.pad("md"), 0)
    )
    kit.entry(body, app.tesseract_path_var, placeholder="Auto-detected if installed").grid(
        row=4, column=1, columnspan=2, sticky="ew", pady=(kit.pad("md"), 0)
    )
    tesseract_actions = kit.button_bar(body, columns=2)
    tesseract_actions.grid(row=5, column=1, columnspan=2, sticky="w", pady=(kit.pad("sm"), 0))
    tesseract_actions.button("Browse", app._browse_tesseract_path)
    tesseract_actions.button("Test", app._test_tesseract_path)

    kit.label(body, "Active route", variant="subheading").grid(
        row=6, column=0, sticky="w", pady=(kit.pad("lg"), 0)
    )
    kit.label(
        body, "", variant="success", textvariable=app.active_engine_var, wraplength=420, justify="left", anchor="w"
    ).grid(row=6, column=1, columnspan=2, sticky="ew", pady=(kit.pad("lg"), 0))

    for row, variable in enumerate((app.dependency_var, app.route_preview_var), start=7):
        kit.label(
            body, "", variant="small", textvariable=variable, wraplength=420, justify="left", anchor="w"
        ).grid(row=row, column=0, columnspan=3, sticky="ew", pady=(kit.pad("sm"), 0))


def _build_files(app, kit, content) -> None:
    import app as app_module
    from app_state import APP_STATE_PATH

    open_path = app_module.open_path

    card = kit.card(content, title="Files & folders")
    app.settings_files_frame = card
    card.grid(row=1, column=0, sticky="nsew", padx=(0, kit.pad("md")))
    body = card.body
    body.grid_columnconfigure(1, weight=1)
    row = 0

    def path_row(label: str, entry, *buttons):
        """Label, an entry, and a trailing button bar. Returns the entry."""
        nonlocal row
        kit.label(body, label, variant="body").grid(
            row=row, column=0, sticky="w", padx=(0, kit.pad("sm")), pady=(kit.pad("sm"), 0)
        )
        entry.grid(row=row, column=1, sticky="ew", pady=(kit.pad("sm"), 0))
        if buttons:
            bar = kit.button_bar(body, columns=3)
            bar.grid(row=row, column=2, sticky="w", padx=(kit.pad("sm"), 0), pady=(kit.pad("sm"), 0))
            for text, command, variant in buttons:
                bar.button(text, command, variant=variant)
        row += 1
        return entry

    app.state_path_entry = path_row(
        "State file", _readonly(kit, body, APP_STATE_PATH),
        ("Open folder", lambda: open_path(APP_STATE_PATH.parent), "secondary"),
    )
    app.notes_path_entry = path_row(
        "Footer notes", _readonly(kit, body, app.notes_path),
        ("Open", lambda: open_path(app.notes_path), "secondary"),
    )
    app.profile_path_entry = path_row(
        "About profile", _readonly(kit, body, app.about_profile_path),
        ("Edit JSON", lambda: open_path(app.about_profile_path), "secondary"),
    )
    app.build_notes_entry = path_row(
        "Build notes", _readonly(kit, body, app.build_notes_path),
        ("Open", lambda: open_path(app.build_notes_path), "secondary"),
    )
    app.shortcuts_path_entry = path_row(
        "Shortcut guide", _readonly(kit, body, app.shortcuts_path),
        ("Open", lambda: open_path(app.shortcuts_path), "secondary"),
    )
    app.static_about_entry = path_row(
        "Installer About", _readonly(kit, body, app.static_about_profile_path),
        ("Open", lambda: open_path(app.static_about_profile_path), "secondary"),
    )
    app.splash_path_entry = path_row(
        "Splash GIF", kit.entry(body, app.splash_gif_path_var),
        ("Browse", app._browse_splash_gif_path, "secondary"),
        ("Preview", app._preview_splash, "secondary"),
        ("Reset", app._clear_splash_gif_path, "ghost"),
    )
    path_row(
        "Install date", kit.entry(body, app.install_date_var),
        ("State folder", lambda: open_path(APP_STATE_PATH.parent), "ghost"),
    )
    path_row(
        "Last update check", kit.entry(body, app.last_update_check_var),
        ("Check now", app._check_for_updates_placeholder, "secondary"),
    )
    path_row(
        "Last state backup", kit.entry(body, app.last_state_backup_var),
        ("Open folder", app._open_state_backup_dir, "secondary"),
        ("Restore", app._restore_latest_state_backup_action, "secondary"),
    )

    kit.separator(body).grid(row=row, column=0, columnspan=3, sticky="ew", pady=kit.pad("md"))
    row += 1

    path_row(
        "Link cache", kit.entry(body, app.link_cache_dir_var),
        ("Open", app._open_link_cache_dir, "secondary"),
        ("Clear", app._clear_link_cache, "secondary"),
        ("Prune", app._prune_link_cache, "ghost"),
    )
    kit.label(
        body, "", variant="small", textvariable=app.link_cache_summary_var, wraplength=420, justify="left", anchor="w"
    ).grid(row=row, column=0, columnspan=3, sticky="ew", pady=(kit.pad("sm"), 0))
    row += 1

    numbers = kit.frame(body, variant="transparent")
    numbers.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(kit.pad("sm"), 0))
    for index, (label, variable) in enumerate(
        (
            ("Keep days", app.link_cache_max_age_var),
            ("Size cap (MB)", app.link_cache_max_size_var),
            ("Timeout (s)", app.link_timeout_var),
        )
    ):
        kit.label(numbers, label, variant="body").grid(row=0, column=index * 2, padx=(0 if index == 0 else kit.pad("md"), kit.pad("sm")))
        kit.entry(numbers, variable, width=80).grid(row=0, column=index * 2 + 1)
    row += 1

    kit.check(body, "Keep downloaded link files in the cache", app.link_keep_downloads_var).grid(
        row=row, column=0, columnspan=3, sticky="w", pady=(kit.pad("sm"), 0)
    )
    row += 1
    kit.button(body, "Refresh cache stats", app._refresh_link_cache_summary, variant="ghost").grid(
        row=row, column=0, columnspan=3, sticky="w", pady=(kit.pad("sm"), 0)
    )
    row += 1

    path_row(
        "Activity reports", kit.entry(body, app.activity_report_dir_var),
        ("Open", app._open_activity_report_dir, "secondary"),
        ("Export", app._export_activity_report_action, "secondary"),
    )
    path_row(
        "Support bundles", kit.entry(body, app.support_bundle_dir_var),
        ("Open", app._open_support_bundle_dir, "secondary"),
        ("Bundle", app._export_support_bundle_action, "secondary"),
    )

    kit.separator(body).grid(row=row, column=0, columnspan=3, sticky="ew", pady=kit.pad("md"))
    row += 1

    kit.check(
        body,
        "Enable optional remote asset refresh",
        app.asset_remote_enabled_var,
        command=app._refresh_asset_status,
    ).grid(row=row, column=0, columnspan=3, sticky="w")
    row += 1

    app.asset_config_entry = path_row(
        "Asset config", _readonly(kit, body, app.asset_config_path),
        ("Open", lambda: open_path(app.asset_config_path), "secondary"),
    )
    path_row(
        "Header image", kit.entry(body, app.header_gif_path_var),
        ("Browse", app._browse_header_gif_path, "secondary"),
        ("Reset", app._clear_header_gif_path, "ghost"),
    )
    path_row(
        "Header URL", kit.entry(body, app.header_gif_url_var),
        ("Refresh assets", app._refresh_remote_assets, "secondary"),
    )
    path_row(
        "Splash URL", kit.entry(body, app.splash_gif_url_var),
        ("Preview", app._preview_splash, "secondary"),
    )
    path_row(
        "About image URL", kit.entry(body, app.about_image_url_var),
        ("Reload About", app._refresh_about_profile, "secondary"),
    )
    path_row(
        "Profile JSON URL", kit.entry(body, app.profile_json_url_var),
        ("Pull now", app._pull_remote_profile, "secondary"),
    )
    path_row(
        "Asset cache", kit.entry(body, app.asset_cache_dir_var),
        ("Open", app._open_asset_cache_dir, "secondary"),
        ("Clear", app._clear_asset_cache_action, "ghost"),
    )

    asset_numbers = kit.frame(body, variant="transparent")
    asset_numbers.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(kit.pad("sm"), 0))
    kit.label(asset_numbers, "Timeout (s)", variant="body").grid(row=0, column=0, padx=(0, kit.pad("sm")))
    kit.entry(asset_numbers, app.asset_timeout_var, width=80).grid(row=0, column=1)
    kit.label(asset_numbers, "Refresh window (h)", variant="body").grid(row=0, column=2, padx=(kit.pad("md"), kit.pad("sm")))
    kit.entry(asset_numbers, app.asset_refresh_hours_var, width=80).grid(row=0, column=3)
    row += 1

    kit.label(
        body, "", variant="small", textvariable=app.asset_status_var, wraplength=420, justify="left", anchor="w"
    ).grid(row=row, column=0, columnspan=3, sticky="ew", pady=(kit.pad("sm"), 0))
    row += 1
    kit.labelled(
        body, "About remote assets",
        "Installer builds always ship local fallback images, so remote assets are "
        "optional and never break offline runs.",
        variant="small",
    ).grid(row=row, column=0, columnspan=3, sticky="w", pady=(kit.pad("sm"), 0))


def _build_quick_actions(app, kit, content) -> None:
    import app as app_module

    open_path = app_module.open_path

    card = kit.card(content, title="Quick actions")
    app.settings_actions_frame = card
    card.grid(row=1, column=1, sticky="nsew")
    body = card.body
    body.grid_columnconfigure(0, weight=1)

    quick_buttons = [
        ("Quick actions palette", app._open_command_palette),
        ("Open Build Center", app._open_build_center_window),
        ("Open SMTP delivery", app._open_smtp_window),
        ("Edit About profile", app._open_about_editor_window),
        ("Open shortcut guide", app._open_shortcuts_window),
        ("Open footer notes", app._open_notes_window),
        ("Preview splash", app._preview_splash),
        ("Reset login reminder", app._reset_login_popup_state),
        ("Refresh remote assets", app._refresh_remote_assets),
        ("Pull remote profile", app._pull_remote_profile),
        ("Export settings snapshot", app._export_state_snapshot_action),
        ("Import settings snapshot", app._import_state_snapshot_action),
        ("Export diagnostics report", app._export_diagnostics_report_action),
        ("Export activity report", app._export_activity_report_action),
        ("Export support bundle", app._export_support_bundle_action),
        ("Export app logs", app._export_current_logs_action),
        ("Send latest outputs", app._send_last_outputs_via_smtp),
        ("Save EML draft", app._save_eml_draft_for_last_outputs),
        ("Check for updates", app._check_for_updates_placeholder),
        ("Open build notes", lambda: open_path(app.build_notes_path)),
        ("Refresh history view", app._refresh_history_views),
        ("Clear local history", app._clear_history),
    ]
    for index, (label, command) in enumerate(quick_buttons):
        variant = "danger" if label.startswith("Clear") else "secondary"
        kit.button(body, label, command, variant=variant).grid(
            row=index, column=0, sticky="ew", pady=(0 if index == 0 else kit.pad("sm"), 0)
        )

    kit.label(
        body, "", variant="small", textvariable=app.smtp_status_var, wraplength=300, justify="left", anchor="w"
    ).grid(row=len(quick_buttons), column=0, sticky="ew", pady=(kit.pad("md"), 0))
