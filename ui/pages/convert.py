"""Convert page: mode picker, image-folder flow, input queue, link ingest, options, log."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path

from ui.dnd import register_drop_target


def build(app) -> None:
    import app as app_module  # constants live on the module, not the instance

    kit = app.kit
    page = app._create_page_frame("convert", scrollable=True)
    page.grid_columnconfigure(0, weight=1)
    page.grid_rowconfigure(3, weight=1)

    _build_mode_card(app, kit, page)
    _build_image_folder_card(app, kit, page)
    _build_center(app, kit, page)
    _build_log_card(app, kit, page)


def _build_mode_card(app, kit, page) -> None:
    import app as app_module

    card = kit.card(page, title="Conversion mode")  # help lives on the mode ⓘ below
    card.grid(row=0, column=0, sticky="ew")
    body = card.body
    body.grid_columnconfigure(1, weight=1)

    kit.label(body, "Mode", variant="body").grid(row=0, column=0, sticky="w", padx=(0, kit.pad("md")))
    app.mode_combo = kit.combo(body, app_module.MODE_ORDER, app.mode_var)
    app.mode_combo.grid(row=0, column=1, sticky="ew")
    app.mode_combo.bind("<<ComboboxSelected>>", app._on_mode_changed)
    kit.button(body, "Refresh dependencies", app._refresh_dependency_status, variant="ghost").grid(
        row=0, column=2, padx=(kit.pad("sm"), 0)
    )

    for row, variable in enumerate(
        (app.mode_help_var, app.supported_var, app.route_preview_var), start=1
    ):
        kit.label(
            body,
            "",
            variant="small",
            textvariable=variable,
            wraplength=900,
            justify="left",
            anchor="w",
        ).grid(row=row, column=0, columnspan=3, sticky="ew", pady=(kit.pad("sm"), 0))


def _build_image_folder_card(app, kit, page) -> None:
    from image_folder_pdf import (
        IMAGE_FOLDER_SCOPE_LABELS,
        IMAGE_FOLDER_SCOPE_OPTIONS,
        IMAGE_FOLDER_SORT_LABELS,
        IMAGE_FOLDER_SORT_OPTIONS,
    )

    card = kit.card(
        page,
        title="Image folder → PDF",
        subtitle="Turn a folder of images into one PDF, or one PDF per subfolder.",
    )
    card.grid(row=1, column=0, sticky="ew", pady=(kit.pad("lg"), 0))
    body = card.body
    body.grid_columnconfigure(1, weight=1)

    kit.label(body, "Folder", variant="body").grid(row=0, column=0, sticky="w", padx=(0, kit.pad("md")))
    app.image_folder_source_entry = kit.entry(
        body, app.image_folder_source_var, placeholder="Pick a folder of images"
    )
    app.image_folder_source_entry.grid(row=0, column=1, sticky="ew")
    register_drop_target(
        app.image_folder_source_entry,
        lambda paths: app.image_folder_source_var.set(str(paths[0])),
    )

    folder_actions = kit.button_bar(body, columns=4)
    folder_actions.grid(row=0, column=2, sticky="e", padx=(kit.pad("sm"), 0))
    folder_actions.button("Browse", app._browse_image_folder_source)
    folder_actions.button("Preview", app._preview_image_folder_summary)
    folder_actions.button("Load queue", app._load_image_folder_to_queue)
    folder_actions.button("Run now", app._start_image_folder_pdf_workflow, variant="primary")

    options = kit.frame(body, variant="transparent")
    options.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(kit.pad("md"), 0))
    options.grid_columnconfigure(1, weight=1)
    options.grid_columnconfigure(3, weight=1)

    kit.label(options, "Name", variant="body").grid(row=0, column=0, sticky="w", padx=(0, kit.pad("sm")))
    kit.entry(options, app.image_folder_output_name_var).grid(row=0, column=1, sticky="ew")
    kit.label(options, "Sort", variant="body").grid(row=0, column=2, sticky="w", padx=(kit.pad("lg"), kit.pad("sm")))
    app.image_folder_sort_combo = kit.combo(
        options,
        [IMAGE_FOLDER_SORT_LABELS.get(item, item) for item in IMAGE_FOLDER_SORT_OPTIONS],
        app.image_folder_sort_var,
    )
    app.image_folder_sort_combo.grid(row=0, column=3, sticky="ew")

    kit.label(options, "Scope", variant="body").grid(
        row=1, column=0, sticky="w", padx=(0, kit.pad("sm")), pady=(kit.pad("sm"), 0)
    )
    app.image_folder_scope_combo = kit.combo(
        options,
        [IMAGE_FOLDER_SCOPE_LABELS.get(item, item) for item in IMAGE_FOLDER_SCOPE_OPTIONS],
        app.image_folder_scope_var,
    )
    app.image_folder_scope_combo.grid(row=1, column=1, sticky="ew", pady=(kit.pad("sm"), 0))

    checks = kit.frame(options, variant="transparent")
    checks.grid(row=1, column=2, columnspan=2, sticky="w", padx=(kit.pad("lg"), 0), pady=(kit.pad("sm"), 0))
    kit.check(checks, "Recursive", app.image_folder_recursive_var).grid(row=0, column=0)
    kit.check(checks, "Include hidden", app.image_folder_include_hidden_var).grid(
        row=0, column=1, padx=(kit.pad("md"), 0)
    )

    kit.label(
        body,
        "",
        variant="small",
        textvariable=app.image_folder_summary_var,
        wraplength=800,
        justify="left",
        anchor="w",
    ).grid(row=2, column=0, columnspan=3, sticky="ew", pady=(kit.pad("sm"), 0))


def _build_center(app, kit, page) -> None:
    center = kit.frame(page, variant="transparent")
    app.convert_center = center
    center.grid(row=2, column=0, sticky="nsew", pady=kit.pad("lg"))
    center.grid_columnconfigure(0, weight=3)
    center.grid_columnconfigure(1, weight=2)
    center.grid_rowconfigure(0, weight=1)

    _build_inputs(app, kit, center)
    _build_options(app, kit, center)


def _build_inputs(app, kit, center) -> None:
    inputs = kit.card(center, title="Input files")
    app.convert_inputs_frame = inputs
    inputs.grid(row=0, column=0, sticky="nsew", padx=(0, kit.pad("md")))
    body = inputs.body
    body.grid_rowconfigure(1, weight=1)
    body.grid_rowconfigure(3, weight=1)

    kit.labelled(
        body,
        "Files to convert",
        "Pick a conversion mode first — adding a folder only picks up the file types "
        "that mode supports.\n\nYou can also drag files straight onto the list.",
        variant="subheading",
    ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, kit.pad("sm")))

    app.file_listbox = kit.listbox(body, selectmode=tk.EXTENDED, height=9)
    app.file_listbox.grid(row=1, column=0, sticky="nsew")
    app.file_listbox.bind("<Double-1>", lambda _event: app._preview_selected_inputs())
    register_drop_target(app.file_listbox, lambda paths: app._append_files([Path(p) for p in paths]))

    side = kit.frame(body, variant="transparent")
    side.grid(row=1, column=1, sticky="ns", padx=(kit.pad("md"), 0))
    for index, (text, command, variant) in enumerate(
        [
            ("Add files", app._add_files, "secondary"),
            ("Add folder", app._add_folder, "secondary"),
            ("Preview selected", app._preview_selected_inputs, "secondary"),
            ("Remove selected", app._remove_selected, "secondary"),
            ("Clear all", app._clear_inputs, "ghost"),
        ]
    ):
        kit.button(side, text, command, variant=variant, width=150).grid(
            row=index, column=0, sticky="ew", pady=(0 if index == 0 else kit.pad("sm"), 0)
        )

    _build_links(app, kit, body)


def _build_links(app, kit, parent) -> None:
    links = kit.card(parent, title="Online links", subtitle="Paste URLs to download and queue them.")
    links.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(kit.pad("lg"), 0))
    body = links.body
    body.grid_rowconfigure(2, weight=1)
    body.grid_rowconfigure(4, weight=1)

    kit.label(
        body,
        "",
        variant="small",
        textvariable=app.link_status_summary_var,
        wraplength=520,
        justify="left",
        anchor="w",
    ).grid(row=0, column=0, sticky="ew")

    actions = kit.button_bar(body, columns=5)
    actions.grid(row=1, column=0, sticky="ew", pady=(kit.pad("sm"), kit.pad("sm")))
    app.link_paste_button = actions.button("Paste URLs", app._paste_urls_from_clipboard)
    app.link_fetch_button = actions.button("Fetch", app._start_link_fetch, variant="primary")
    app.link_fetch_start_button = actions.button("Fetch + start", lambda: app._start_link_fetch(auto_start=True))
    app.link_retry_button = actions.button("Retry failed", lambda: app._start_link_fetch(retry_failed=True))
    app.link_pause_button = actions.button("Pause", app._pause_link_fetch)
    app.link_resume_button = actions.button("Resume", app._resume_link_fetch)
    app.link_cancel_button = actions.button("Cancel", app._cancel_link_fetch, variant="ghost")
    actions.button("Clear URLs", app._clear_link_urls, variant="ghost")
    actions.button("Load recent", app._load_recent_links, variant="ghost")
    actions.button("Open cache", app._open_link_cache_dir, variant="ghost")

    app.link_input_text = kit.textbox(body, mono=False, height=110)
    app.link_input_text.grid(row=2, column=0, sticky="nsew")

    kit.label(
        body,
        "",
        variant="small",
        textvariable=app.link_recent_summary_var,
        wraplength=640,
        justify="left",
        anchor="w",
    ).grid(row=3, column=0, sticky="ew", pady=(kit.pad("sm"), kit.pad("sm")))

    app.link_status_tree = kit.treeview(
        body, columns=("status", "url", "file", "detail"), show="headings", height=5
    )
    for column, title, width in (
        ("status", "Status", 110),
        ("url", "URL", 260),
        ("file", "Local file", 200),
        ("detail", "Detail", 120),
    ):
        app.link_status_tree.heading(column, text=title)
        app.link_status_tree.column(column, width=width, anchor="w")
    app.link_status_tree.grid(row=4, column=0, sticky="nsew")
    app.link_status_tree.bind("<<TreeviewSelect>>", app._on_link_status_selected)

    kit.label(body, "", variant="tiny", textvariable=app.link_fetch_count_var).grid(
        row=5, column=0, sticky="w", pady=(kit.pad("sm"), 0)
    )
    app.after_idle(app._update_link_fetch_buttons)


def _build_options(app, kit, center) -> None:
    options = kit.card(center, title="Output & options")
    app.convert_options_frame = options
    options.grid(row=0, column=1, sticky="nsew")
    body = options.body
    body.grid_columnconfigure(1, weight=1)

    kit.label(body, "Output folder", variant="body").grid(row=0, column=0, sticky="w", padx=(0, kit.pad("sm")))
    app.output_entry = kit.entry(body, app.output_dir_var)
    app.output_entry.grid(row=0, column=1, sticky="ew")

    output_actions = kit.button_bar(body, columns=2)
    output_actions.grid(row=1, column=1, sticky="e", pady=(kit.pad("sm"), 0))
    output_actions.button("Browse", app._browse_output_dir)
    output_actions.button("Open folder", app._open_output_folder, variant="ghost")

    app.merge_check = kit.check(body, "Merge into one PDF", app.merge_var, command=app._update_mode_controls)
    app.merge_check.grid(row=2, column=0, columnspan=2, sticky="w", pady=(kit.pad("md"), 0))

    kit.label(body, "Merged name", variant="body").grid(
        row=3, column=0, sticky="w", padx=(0, kit.pad("sm")), pady=(kit.pad("sm"), 0)
    )
    app.output_name_entry = kit.entry(body, app.output_name_var)
    app.output_name_entry.grid(row=3, column=1, sticky="ew", pady=(kit.pad("sm"), 0))

    app.recursive_check = kit.check(body, "Scan folders recursively", app.recursive_var)
    app.recursive_check.grid(row=4, column=0, columnspan=2, sticky="w", pady=(kit.pad("md"), 0))

    kit.label(body, "Image format", variant="body").grid(
        row=5, column=0, sticky="w", padx=(0, kit.pad("sm")), pady=(kit.pad("sm"), 0)
    )
    app.image_format_combo = kit.combo(body, ["png", "jpg"], app.image_format_var)
    app.image_format_combo.grid(row=5, column=1, sticky="w", pady=(kit.pad("sm"), 0))

    kit.labelled(
        body, "Image scale",
        "How large exported images are.\n\n2.0 is a good default. Higher means "
        "sharper but bigger files; lower means smaller and faster.",
    ).grid(row=6, column=0, sticky="w", padx=(0, kit.pad("sm")), pady=(kit.pad("sm"), 0))
    app.image_scale_entry = kit.entry(body, app.image_scale_var, width=100)
    app.image_scale_entry.grid(row=6, column=1, sticky="w", pady=(kit.pad("sm"), 0))



    kit.label(body, "Dependencies", variant="subheading").grid(
        row=8, column=0, columnspan=2, sticky="w", pady=(kit.pad("lg"), 0)
    )
    kit.label(
        body,
        "",
        variant="small",
        textvariable=app.dependency_var,
        wraplength=320,
        justify="left",
        anchor="w",
    ).grid(row=9, column=0, columnspan=2, sticky="ew", pady=(kit.pad("xs"), 0))


def _build_log_card(app, kit, page) -> None:
    card = kit.card(page, title="Activity")
    card.grid(row=3, column=0, sticky="nsew")
    body = card.body
    body.grid_rowconfigure(1, weight=1)

    top = kit.frame(body, variant="transparent")
    top.grid(row=0, column=0, sticky="ew", pady=(0, kit.pad("sm")))
    top.grid_columnconfigure(0, weight=1)
    app.progress = kit.progress(top)
    app.progress.grid(row=0, column=0, sticky="ew", padx=(0, kit.pad("md")))
    kit.button(top, "Clear log", app._clear_log, variant="ghost").grid(row=0, column=1, padx=(0, kit.pad("sm")))
    app.start_button = kit.button(top, "Start conversion", app._start_conversion, variant="primary")
    app.start_button.grid(row=0, column=2)

    app.log_text = kit.textbox(body, height=240)
    app.log_text.grid(row=1, column=0, sticky="nsew")
    app.log_text.configure(state="disabled")
