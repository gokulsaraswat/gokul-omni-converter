"""Automation page: saved presets, the watch folder, and output bundling."""

from __future__ import annotations


def build(app) -> None:
    kit = app.kit
    page = app._create_page_frame("automation", scrollable=True)
    page.grid_columnconfigure(0, weight=1)
    page.grid_rowconfigure(1, weight=1)
    app.preset_item_ids: dict[str, dict[str, object]] = {}

    kit.page_header(
        page,
        "Automation",
        "Reusable presets and a watch folder that converts new files as they land, "
        "then bundles or reports the results.",
    ).grid(row=0, column=0, sticky="ew", pady=(0, kit.pad("lg")))

    content = kit.frame(page, variant="transparent")
    # Deliberately NOT settings_content_frame: the settings page owns that name for
    # the responsive breakpoints, and both pages using it made the layout depend on
    # build order.
    app.automation_content_frame = content
    content.grid(row=1, column=0, sticky="nsew")
    content.grid_columnconfigure(0, weight=1)
    content.grid_columnconfigure(1, weight=1)
    content.grid_rowconfigure(0, weight=1)
    content.grid_rowconfigure(1, weight=1)

    _build_presets(app, kit, content)
    _build_watch(app, kit, content)
    _build_lower(app, kit, content)

    app._restore_automation_log()
    app._refresh_presets_view()
    app._refresh_automation_page()


def _build_presets(app, kit, content) -> None:
    card = kit.card(
        content,
        title="Saved presets",
        subtitle="Save the current Convert settings under a name, then re-apply them later.",
    )
    card.grid(row=0, column=0, sticky="nsew", padx=(0, kit.pad("md")), pady=(0, kit.pad("md")))
    body = card.body
    body.grid_rowconfigure(1, weight=1)

    name_row = kit.frame(body, variant="transparent")
    name_row.grid(row=0, column=0, sticky="ew", pady=(0, kit.pad("md")))
    name_row.grid_columnconfigure(1, weight=1)
    kit.label(name_row, "Name", variant="body").grid(row=0, column=0, padx=(0, kit.pad("sm")))
    kit.entry(name_row, app.preset_name_var, placeholder="e.g. Invoices → PDF").grid(row=0, column=1, sticky="ew")
    kit.button(name_row, "Save current", app._save_current_preset, variant="primary").grid(
        row=0, column=2, padx=(kit.pad("sm"), 0)
    )

    app.preset_tree = kit.treeview(
        body, columns=("favorite", "name", "mode", "engine", "merge"), show="headings", height=9
    )
    for column, title, width, anchor in (
        ("favorite", "★", 44, "center"),
        ("name", "Name", 140, "w"),
        ("mode", "Mode", 190, "w"),
        ("engine", "Engine", 90, "center"),
        ("merge", "Merge", 70, "center"),
    ):
        app.preset_tree.heading(column, text=title)
        app.preset_tree.column(column, width=width, anchor=anchor)
    app.preset_tree.grid(row=1, column=0, sticky="nsew")
    app.preset_tree.bind("<<TreeviewSelect>>", app._on_preset_selected)

    actions = kit.button_bar(body, columns=3)
    actions.grid(row=2, column=0, sticky="ew", pady=(kit.pad("md"), 0))
    actions.button("Apply to Convert", app._apply_selected_preset, variant="primary")
    actions.button("★ Favourite", app._toggle_selected_preset_favorite)
    actions.button("Delete", app._delete_selected_preset, variant="danger")
    actions.button("Export JSON", app._export_presets, variant="ghost")
    actions.button("Import JSON", app._import_presets, variant="ghost")


def _build_watch(app, kit, content) -> None:
    from converter_core import ENGINE_ORDER

    import app as app_module

    card = kit.card(content, title="Watch folder")
    card.grid(row=0, column=1, sticky="nsew", pady=(0, kit.pad("md")))
    body = card.body
    body.grid_columnconfigure(1, weight=1)

    row = kit.field_row(
        body,
        0,
        "Source folder",
        kit.entry(body, app.watch_source_dir_var, placeholder="Folder to watch"),
        trailing=kit.button(body, "Browse", app._browse_watch_source_dir, variant="secondary", width=90),
    )
    row = kit.field_row(
        body,
        row,
        "Output folder",
        kit.entry(body, app.watch_output_dir_var),
        trailing=kit.button(body, "Browse", app._browse_watch_output_dir, variant="secondary", width=90),
    )
    row = kit.field_row(body, row, "Mode", kit.combo(body, app_module.MODE_ORDER, app.watch_mode_var))
    row = kit.field_row(body, row, "Engine", kit.combo(body, list(ENGINE_ORDER), app.watch_engine_var))
    row = kit.field_row(body, row, "Merged name", kit.entry(body, app.watch_output_name_var))
    row = kit.field_row(body, row, "Archive folder", kit.entry(body, app.watch_archive_dir_var),
                        trailing=kit.button(body, "Browse", app._browse_watch_archive_dir, variant="secondary", width=90))

    interval = kit.frame(body, variant="transparent")
    interval.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(kit.pad("sm"), 0))
    kit.check(interval, "Recursive scan", app.watch_recursive_var).grid(row=0, column=0, sticky="w")
    kit.label(interval, "Interval (s)", variant="body").grid(row=0, column=1, padx=(kit.pad("lg"), kit.pad("sm")))
    kit.entry(interval, app.watch_interval_var, width=80).grid(row=0, column=2)
    row += 1

    for text, variable in (
        ("Merge outputs into one PDF when supported", app.watch_merge_var),
        ("Archive processed source files", app.watch_archive_var),
        ("Create a ZIP bundle after each run", app.watch_zip_var),
        ("Create a text report after each run", app.watch_report_var),
        ("Open a mail draft after each run", app.watch_mail_var),
        ("Skip existing files when the watcher starts", app.watch_skip_existing_var),
    ):
        kit.check(body, text, variable).grid(row=row, column=0, columnspan=3, sticky="w", pady=(kit.pad("xs"), 0))
        row += 1

    actions = kit.button_bar(body, columns=2)
    actions.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(kit.pad("md"), 0))
    actions.button("Save config", app._save_watch_config)
    actions.button("Scan now", app._scan_watch_folder_now)
    actions.button("Start watcher", app._start_watch_automation, variant="primary")
    actions.button("Stop watcher", app._stop_watch_automation, variant="danger")
    row += 1

    kit.label(
        body,
        "",
        variant="small",
        textvariable=app.watch_status_var,
        wraplength=420,
        justify="left",
        anchor="w",
    ).grid(row=row, column=0, columnspan=3, sticky="ew", pady=(kit.pad("md"), 0))


def _build_lower(app, kit, content) -> None:
    lower = kit.frame(content, variant="transparent")
    lower.grid(row=1, column=0, columnspan=2, sticky="nsew")
    lower.grid_columnconfigure(0, weight=1)
    lower.grid_columnconfigure(1, weight=1)
    lower.grid_rowconfigure(0, weight=1)

    share = kit.card(lower, title="Bundles & reports")
    share.grid(row=0, column=0, sticky="nsew", padx=(0, kit.pad("md")))
    body = share.body
    body.grid_columnconfigure(1, weight=1)

    kit.label(
        body,
        "",
        variant="small",
        textvariable=app.watch_summary_var,
        wraplength=420,
        justify="left",
        anchor="w",
    ).grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, kit.pad("md")))

    row = kit.field_row(body, 1, "Bundle name", kit.entry(body, app.share_bundle_name_var))
    row = kit.field_row(body, row, "Report name", kit.entry(body, app.share_report_name_var))

    actions = kit.button_bar(body, columns=2)
    actions.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(kit.pad("md"), 0))
    actions.button("ZIP last outputs", app._create_zip_bundle_for_last_outputs, variant="primary")
    actions.button("Export last report", app._export_last_run_report)
    actions.button("Mail draft", app._open_mail_draft_for_last_outputs)
    actions.button("Save EML", app._save_eml_draft_for_last_outputs)
    actions.button("Open last output folder", app._open_last_output_folder, variant="ghost")
    actions.button("Send latest outputs", app._send_last_outputs_via_smtp, variant="ghost")
    actions.button("Reset seen cache", app._clear_watch_seen_cache, variant="ghost")
    actions.button("Refresh page", app._refresh_automation_page, variant="ghost")

    log = kit.card(lower, title="Automation events")
    log.grid(row=0, column=1, sticky="nsew")
    log.body.grid_rowconfigure(0, weight=1)
    app.automation_log_text = kit.textbox(log.body, height=300)
    app.automation_log_text.grid(row=0, column=0, sticky="nsew")
    app.automation_log_text.configure(state="disabled")
