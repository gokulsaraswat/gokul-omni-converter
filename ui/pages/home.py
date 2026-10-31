"""Home / quick-start page.

Assigns its widgets onto ``app`` under the same attribute names the pre-2.4.0
ttk build used, so every handler in app.py keeps working unchanged.
"""

from __future__ import annotations


def build(app) -> None:
    kit = app.kit
    page = app._create_page_frame("home", scrollable=True)
    page.grid_columnconfigure(0, weight=1)
    page.grid_rowconfigure(2, weight=1)

    # -- hero ---------------------------------------------------------------
    hero = kit.card(page)
    hero.grid(row=0, column=0, sticky="ew")
    hero.grid_columnconfigure(0, weight=1)

    hero_text = kit.frame(hero, variant="transparent")
    hero_text.grid(row=0, column=0, sticky="w", padx=kit.pad("lg"), pady=kit.pad("lg"))
    kit.label(hero_text, "Quick start", variant="title").grid(row=0, column=0, sticky="w")
    kit.label(
        hero_text,
        "Add files, a folder, or a link — then run the workflow you need.",
        variant="muted",
        wraplength=520,
        justify="left",
        anchor="w",
    ).grid(row=1, column=0, sticky="w", pady=(kit.pad("xs"), 0))

    actions = kit.button_bar(hero, columns=3)
    actions.grid(row=0, column=1, sticky="e", padx=kit.pad("lg"), pady=kit.pad("lg"))
    actions.button("Add files", app._add_files)
    actions.button("Add folder", app._add_folder)
    actions.button("Image folder → PDF", app._focus_image_folder_flow)
    actions.button("Paste links", app._focus_link_input)
    actions.button("Quick actions", app._open_command_palette, variant="ghost")
    actions.button(
        "Start batch",
        lambda: (app._show_page("convert"), app._start_conversion()),
        variant="primary",
    )

    # -- metric strip -------------------------------------------------------
    metrics = kit.frame(page, variant="transparent")
    app.home_metrics_frame = metrics
    metrics.grid(row=1, column=0, sticky="ew", pady=kit.pad("lg"))
    for column in range(4):
        metrics.grid_columnconfigure(column, weight=1)

    app._create_metric_card(metrics, 0, "Inputs", app.home_selected_count_var)
    app._create_metric_card(metrics, 1, "Mode", app.home_mode_var)
    app._create_metric_card(metrics, 2, "Engine", app.home_engine_var)
    app._create_metric_card(metrics, 3, "Output", app.home_output_var)

    # -- lower split --------------------------------------------------------
    lower = kit.frame(page, variant="transparent")
    app.home_lower_frame = lower
    lower.grid(row=2, column=0, sticky="nsew")
    lower.grid_columnconfigure(0, weight=2)
    lower.grid_columnconfigure(1, weight=1)
    lower.grid_rowconfigure(0, weight=1)

    history_card = kit.card(lower, title="Recent jobs")
    app.home_history_card = history_card
    history_card.grid(row=0, column=0, sticky="nsew", padx=(0, kit.pad("md")))
    history_card.body.grid_rowconfigure(0, weight=1)

    app.home_history_tree = kit.treeview(
        history_card.body,
        columns=("time", "status", "mode", "files"),
        show="headings",
        height=8,
    )
    for column, title, width in (
        ("time", "Time", 168),
        ("status", "Status", 104),
        ("mode", "Mode", 320),
        ("files", "Files", 76),
    ):
        app.home_history_tree.heading(column, text=title)
        app.home_history_tree.column(column, width=width, anchor="w")
    app.home_history_tree.grid(row=0, column=0, sticky="nsew")
    app.home_history_tree.bind("<<TreeviewSelect>>", app._on_home_history_selected)

    dep_card = kit.card(lower, title="Status & shortcuts")
    app.home_dep_card = dep_card
    dep_card.grid(row=0, column=1, sticky="nsew")
    body = dep_card.body

    app.home_dependency_label = kit.label(
        body,
        "",
        variant="small",
        textvariable=app.home_dependency_summary_var,
        wraplength=300,
        justify="left",
        anchor="w",
    )
    app.home_dependency_label.grid(row=0, column=0, sticky="ew")

    shortcuts = kit.button_bar(body, columns=2)
    shortcuts.grid(row=1, column=0, sticky="ew", pady=(kit.pad("md"), 0))
    shortcuts.button("PDF Tools", lambda: app._show_page("pdf_tools"))
    shortcuts.button("OCR", lambda: app._show_page("ocr"))
    shortcuts.button("Archive", lambda: app._show_page("archive"))
    shortcuts.button("Rename", lambda: app._show_page("rename_organize"))

    kit.label(
        body,
        "",
        variant="small",
        textvariable=app.home_hint_var,
        wraplength=300,
        justify="left",
        anchor="w",
    ).grid(row=2, column=0, sticky="ew", pady=(kit.pad("md"), 0))

    kit.label(body, "Favourites", variant="subheading").grid(row=3, column=0, sticky="w", pady=(kit.pad("lg"), 0))
    kit.label(
        body,
        "",
        variant="small",
        textvariable=app.favorite_preset_summary_var,
        wraplength=300,
        justify="left",
        anchor="w",
    ).grid(row=4, column=0, sticky="w", pady=(kit.pad("xs"), 0))

    app.home_favorite_presets_frame = kit.frame(body, variant="transparent")
    app.home_favorite_presets_frame.grid(row=5, column=0, sticky="ew", pady=(kit.pad("sm"), 0))
    app.home_favorite_presets_frame.grid_columnconfigure(0, weight=1)
