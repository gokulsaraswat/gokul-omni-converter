"""History page: past jobs, the selected job's detail, recent outputs, and retries."""

from __future__ import annotations

import tkinter as tk


def build(app) -> None:
    kit = app.kit
    page = app._create_page_frame("history", scrollable=True)
    page.grid_columnconfigure(0, weight=2)
    page.grid_columnconfigure(1, weight=1)
    page.grid_rowconfigure(0, weight=1)

    # -- left: the job list -------------------------------------------------
    left = kit.card(
        page,
        title="Recent jobs",
        subtitle="Job summaries are stored locally so you can inspect past runs, "
        "filter them, and load the same settings again.",
    )
    left.grid(row=0, column=0, sticky="nsew", padx=(0, kit.pad("md")))
    body = left.body
    body.grid_rowconfigure(1, weight=1)

    filter_row = kit.frame(body, variant="transparent")
    filter_row.grid(row=0, column=0, sticky="ew", pady=(0, kit.pad("md")))
    filter_row.grid_columnconfigure(1, weight=1)
    kit.label(filter_row, "Filter", variant="body").grid(row=0, column=0, padx=(0, kit.pad("sm")))
    filter_entry = kit.entry(filter_row, app.history_filter_var, placeholder="Search mode, status, or filename")
    filter_entry.grid(row=0, column=1, sticky="ew")
    filter_entry.bind("<KeyRelease>", lambda _event: app._refresh_history_views())
    kit.button(filter_row, "Clear", app._clear_history_filter, variant="ghost").grid(
        row=0, column=2, padx=(kit.pad("sm"), 0)
    )

    app.history_tree = kit.treeview(
        body,
        columns=("time", "status", "mode", "files", "outputs"),
        show="headings",
        height=14,
    )
    for column, title, width in (
        ("time", "Time", 170),
        ("status", "Status", 120),
        ("mode", "Mode", 320),
        ("files", "Files", 70),
        ("outputs", "Outputs", 80),
    ):
        app.history_tree.heading(column, text=title)
        app.history_tree.column(column, width=width, anchor="w")
    app.history_tree.grid(row=1, column=0, sticky="nsew")
    app.history_tree.bind("<<TreeviewSelect>>", app._on_history_selected)

    # -- right: detail + actions -------------------------------------------
    right = kit.card(page, title="Selected job")
    right.grid(row=0, column=1, sticky="nsew")
    detail = right.body
    detail.grid_rowconfigure(1, weight=1)
    detail.grid_rowconfigure(3, weight=1)

    kit.label(
        detail,
        "",
        variant="small",
        textvariable=app.history_detail_var,
        wraplength=320,
        justify="left",
        anchor="w",
    ).grid(row=0, column=0, sticky="ew")

    app.history_details_text = kit.textbox(detail, height=220)
    app.history_details_text.grid(row=1, column=0, sticky="nsew", pady=(kit.pad("sm"), 0))
    app.history_details_text.configure(state="disabled")

    actions = kit.frame(detail, variant="transparent")
    actions.grid(row=2, column=0, sticky="ew", pady=(kit.pad("md"), 0))
    actions.grid_columnconfigure(0, weight=1)
    for index, (text, command, variant) in enumerate(
        [
            ("Load selected settings", app._load_selected_history_settings, "primary"),
            ("Preview selected outputs", app._preview_selected_history_outputs, "secondary"),
            ("Open output folder", app._open_selected_history_output, "secondary"),
            ("Export run report", app._export_selected_history_report, "secondary"),
            ("Export activity report", app._export_activity_report_action, "ghost"),
            ("Export logs", app._export_current_logs_action, "ghost"),
            ("Clear history", app._clear_history, "danger"),
        ]
    ):
        kit.button(actions, text, command, variant=variant).grid(
            row=index, column=0, sticky="ew", pady=(0 if index == 0 else kit.pad("sm"), 0)
        )

    # -- lower: recent outputs and failed jobs ------------------------------
    lower = kit.frame(detail, variant="transparent")
    lower.grid(row=3, column=0, sticky="nsew", pady=(kit.pad("lg"), 0))
    lower.grid_columnconfigure(0, weight=1)
    lower.grid_rowconfigure(1, weight=1)
    lower.grid_rowconfigure(4, weight=1)

    kit.label(lower, "Recent outputs", variant="subheading").grid(row=0, column=0, sticky="w")
    app.recent_outputs_listbox = kit.listbox(lower, selectmode=tk.SINGLE, exportselection=False, height=5)
    app.recent_outputs_listbox.grid(row=1, column=0, sticky="nsew", pady=(kit.pad("xs"), 0))
    app.recent_outputs_listbox.bind("<Double-1>", lambda _event: app._preview_selected_recent_output())

    recent_actions = kit.button_bar(lower, columns=2)
    recent_actions.grid(row=2, column=0, sticky="ew", pady=(kit.pad("sm"), 0))
    recent_actions.button("Preview", app._preview_selected_recent_output)
    recent_actions.button("Open file", app._open_selected_recent_output)
    recent_actions.button("Open folder", app._open_recent_output_folder)
    recent_actions.button("Clear", app._clear_recent_outputs, variant="ghost")

    kit.label(lower, "Failed jobs", variant="subheading").grid(row=3, column=0, sticky="w", pady=(kit.pad("lg"), 0))
    app.failed_jobs_listbox = kit.listbox(lower, selectmode=tk.SINGLE, exportselection=False, height=5)
    app.failed_jobs_listbox.grid(row=4, column=0, sticky="nsew", pady=(kit.pad("xs"), 0))

    failed_actions = kit.button_bar(lower, columns=2)
    failed_actions.grid(row=5, column=0, sticky="ew", pady=(kit.pad("sm"), 0))
    failed_actions.button("Retry selected", app._retry_selected_failed_job, variant="primary")
    failed_actions.button("Remove selected", app._remove_selected_failed_job)
    failed_actions.button("Clear all", app._clear_failed_jobs, variant="ghost")
