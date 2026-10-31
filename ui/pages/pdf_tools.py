"""PDF Tools page.

The options panel exposes every field any tool might need; _update_pdf_tool_controls
in app.py enables and disables them per selected tool, so every label and entry here
keeps the attribute name that handler expects.
"""

from __future__ import annotations

import tkinter as tk
from pathlib import Path

from ui.dnd import register_drop_target


def build(app) -> None:
    kit = app.kit
    page = app._create_page_frame("pdf_tools", scrollable=True)
    page.grid_columnconfigure(0, weight=1)
    page.grid_rowconfigure(2, weight=1)

    _build_tool_card(app, kit, page)
    _build_center(app, kit, page)
    _build_log_card(app, kit, page)


def _build_tool_card(app, kit, page) -> None:
    from converter_core import PDF_TOOL_ORDER

    card = kit.card(
        page,
        title="PDF studio",
        subtitle="Organise, secure, compress, redact, sign, and edit PDFs.",
    )
    card.grid(row=0, column=0, sticky="ew")
    body = card.body
    body.grid_columnconfigure(1, weight=1)

    kit.label(body, "Tool", variant="body").grid(row=0, column=0, sticky="w", padx=(0, kit.pad("md")))
    app.pdf_tool_combo = kit.combo(body, list(PDF_TOOL_ORDER), app.pdf_tool_var)
    app.pdf_tool_combo.grid(row=0, column=1, sticky="ew")
    app.pdf_tool_combo.bind("<<ComboboxSelected>>", app._on_pdf_tool_changed)
    kit.button(body, "Open output folder", app._open_output_folder, variant="ghost").grid(
        row=0, column=2, padx=(kit.pad("sm"), 0)
    )

    for row, variable in enumerate((app.pdf_tool_help_var, app.pdf_tool_hint_var), start=1):
        kit.label(
            body,
            "",
            variant="small",
            textvariable=variable,
            wraplength=900,
            justify="left",
            anchor="w",
        ).grid(row=row, column=0, columnspan=3, sticky="ew", pady=(kit.pad("sm"), 0))


def _build_center(app, kit, page) -> None:
    center = kit.frame(page, variant="transparent")
    center.grid(row=1, column=0, sticky="nsew", pady=kit.pad("lg"))
    center.grid_columnconfigure(0, weight=3)
    center.grid_columnconfigure(1, weight=2)
    center.grid_rowconfigure(0, weight=1)

    # -- input list ---------------------------------------------------------
    inputs = kit.card(
        center,
        title="Selected PDFs",
        subtitle="Merge, batch watermark, and batch edit all follow this list order — "
        "use Move up / Move down when the order matters.",
    )
    inputs.grid(row=0, column=0, sticky="nsew", padx=(0, kit.pad("md")))
    body = inputs.body
    body.grid_rowconfigure(0, weight=1)

    app.pdf_tool_listbox = kit.listbox(body, selectmode=tk.EXTENDED, height=14)
    app.pdf_tool_listbox.grid(row=0, column=0, sticky="nsew")
    register_drop_target(
        app.pdf_tool_listbox,
        lambda paths: app._append_pdf_tool_files([Path(p) for p in paths if Path(p).suffix.lower() == ".pdf"]),
    )

    side = kit.frame(body, variant="transparent")
    side.grid(row=0, column=1, sticky="ns", padx=(kit.pad("md"), 0))
    for index, (text, command, variant) in enumerate(
        [
            ("Add PDFs", app._add_pdf_tool_files, "secondary"),
            ("Add folder", app._add_pdf_tool_folder, "secondary"),
            ("Move up", app._move_pdf_tool_selected_up, "secondary"),
            ("Move down", app._move_pdf_tool_selected_down, "secondary"),
            ("Remove selected", app._remove_pdf_tool_selected, "secondary"),
            ("Clear all", app._clear_pdf_tool_inputs, "ghost"),
        ]
    ):
        kit.button(side, text, command, variant=variant, width=150).grid(
            row=index, column=0, sticky="ew", pady=(0 if index == 0 else kit.pad("sm"), 0)
        )

    _build_options(app, kit, center)


def _build_options(app, kit, center) -> None:
    options = kit.card(center, title="Tool options")
    options.grid(row=0, column=1, sticky="nsew")
    body = options.body
    body.grid_columnconfigure(1, weight=1)
    row = 0

    def text_row(label: str, variable, *, show: str = "", width: int = 0):
        """One label + entry pair. Returns (label_widget, entry_widget)."""
        nonlocal row
        label_widget = kit.label(body, label, variant="body")
        label_widget.grid(row=row, column=0, sticky="w", padx=(0, kit.pad("sm")), pady=(kit.pad("sm"), 0))
        kwargs = {"show": show} if show else {}
        if width:
            kwargs["width"] = width
        entry = kit.entry(body, variable, **kwargs)
        entry.grid(row=row, column=1, columnspan=2, sticky="ew", pady=(kit.pad("sm"), 0))
        row += 1
        return label_widget, entry

    kit.label(body, "Output folder", variant="body").grid(row=row, column=0, sticky="w", padx=(0, kit.pad("sm")))
    kit.entry(body, app.output_dir_var).grid(row=row, column=1, sticky="ew")
    kit.button(body, "Browse", app._browse_output_dir, variant="secondary", width=100).grid(
        row=row, column=2, padx=(kit.pad("sm"), 0)
    )
    row += 1

    _, app.pdf_tool_output_name_entry = text_row("Merged name", app.pdf_tool_output_name_var)
    app.pdf_tool_page_spec_label, app.pdf_tool_page_spec_entry = text_row("Pages / ranges", app.pdf_tool_page_spec_var)
    _, app.pdf_tool_every_n_entry = text_row("Split every N pages", app.pdf_tool_every_n_var, width=110)
    app.pdf_tool_text_label, app.pdf_tool_watermark_text_entry = text_row("Watermark text", app.pdf_tool_watermark_text_var)

    # Watermark image needs a Browse button alongside its entry.
    app.pdf_tool_image_label = kit.label(body, "Watermark image", variant="body")
    app.pdf_tool_image_label.grid(row=row, column=0, sticky="w", padx=(0, kit.pad("sm")), pady=(kit.pad("sm"), 0))
    app.pdf_tool_image_entry = kit.entry(body, app.pdf_tool_watermark_image_var)
    app.pdf_tool_image_entry.grid(row=row, column=1, sticky="ew", pady=(kit.pad("sm"), 0))
    app.pdf_tool_image_browse_button = kit.button(
        body, "Browse", app._browse_pdf_tool_watermark_image, variant="secondary", width=100
    )
    app.pdf_tool_image_browse_button.grid(row=row, column=2, padx=(kit.pad("sm"), 0), pady=(kit.pad("sm"), 0))
    row += 1

    import app as app_module

    app.pdf_tool_position_label = kit.label(body, "Position", variant="body")
    app.pdf_tool_position_label.grid(row=row, column=0, sticky="w", padx=(0, kit.pad("sm")), pady=(kit.pad("sm"), 0))
    app.pdf_tool_position_combo = kit.combo(body, list(app_module.PDF_TOOL_POSITIONS), app.pdf_tool_position_var)
    app.pdf_tool_position_combo.grid(row=row, column=1, sticky="w", pady=(kit.pad("sm"), 0))
    row += 1

    app.pdf_tool_font_size_label, app.pdf_tool_font_size_entry = text_row("Font size", app.pdf_tool_font_size_var, width=110)
    app.pdf_tool_rotation_label, app.pdf_tool_rotation_entry = text_row("Rotation", app.pdf_tool_rotation_var, width=110)
    app.pdf_tool_opacity_label, app.pdf_tool_opacity_entry = text_row("Opacity", app.pdf_tool_opacity_var, width=110)
    app.pdf_tool_image_scale_label, app.pdf_tool_image_scale_entry = text_row(
        "Image scale %", app.pdf_tool_image_scale_var, width=110
    )
    app.pdf_tool_metadata_title_label, app.pdf_tool_metadata_title_entry = text_row(
        "Metadata title", app.pdf_tool_metadata_title_var
    )
    app.pdf_tool_metadata_author_label, app.pdf_tool_metadata_author_entry = text_row(
        "Metadata author", app.pdf_tool_metadata_author_var
    )
    app.pdf_tool_metadata_subject_label, app.pdf_tool_metadata_subject_entry = text_row(
        "Metadata subject", app.pdf_tool_metadata_subject_var
    )
    app.pdf_tool_metadata_keywords_label, app.pdf_tool_metadata_keywords_entry = text_row(
        "Metadata keywords", app.pdf_tool_metadata_keywords_var
    )

    app.pdf_tool_metadata_clear_check = kit.check(
        body, "Clear existing metadata before applying new values", app.pdf_tool_metadata_clear_var
    )
    app.pdf_tool_metadata_clear_check.grid(row=row, column=0, columnspan=3, sticky="w", pady=(kit.pad("md"), 0))
    row += 1

    app.pdf_tool_redact_rect_label, app.pdf_tool_redact_rect_entry = text_row("Area rectangle", app.pdf_tool_redact_rect_var)
    app.pdf_tool_replacement_text_label, app.pdf_tool_replacement_text_entry = text_row(
        "Replacement text", app.pdf_tool_replacement_text_var
    )
    app.pdf_tool_password_label, app.pdf_tool_password_entry = text_row("Password", app.pdf_tool_password_var, show="•")
    app.pdf_tool_owner_password_label, app.pdf_tool_owner_password_entry = text_row(
        "Owner password", app.pdf_tool_owner_password_var, show="•"
    )

    app.pdf_tool_compression_label = kit.label(body, "Compression", variant="body")
    app.pdf_tool_compression_label.grid(row=row, column=0, sticky="w", padx=(0, kit.pad("sm")), pady=(kit.pad("sm"), 0))
    app.pdf_tool_compression_combo = kit.combo(body, ["safe", "balanced", "strong"], app.pdf_tool_compression_profile_var)
    app.pdf_tool_compression_combo.grid(row=row, column=1, sticky="w", pady=(kit.pad("sm"), 0))
    row += 1

    # Per-tool guidance is set by _update_pdf_tool_controls via set_help().
    help_row = kit.frame(body, variant="transparent")
    help_row.grid(row=row, column=0, columnspan=3, sticky="w", pady=(kit.pad("md"), 0))
    kit.label(help_row, "How to enter pages", variant="small").grid(row=0, column=0, sticky="w")
    app.pdf_tool_options_help_label = kit.info(
        help_row,
        "Split into separate PDFs:  1-3; 4-6; 8\n"
        "Extract or remove pages:  1-2, 5, 9-last\n"
        "Reorder (and duplicate):  3,1,2,2\n\n"
        "Opacity runs from 0.01 (barely visible) to 1.0 (solid).",
    )
    app.pdf_tool_options_help_label.grid(row=0, column=1, padx=(kit.pad("xs"), 0))


def _build_log_card(app, kit, page) -> None:
    card = kit.card(page, title="PDF tool activity")
    card.grid(row=2, column=0, sticky="nsew")
    body = card.body
    body.grid_rowconfigure(1, weight=1)

    top = kit.frame(body, variant="transparent")
    top.grid(row=0, column=0, sticky="ew", pady=(0, kit.pad("sm")))
    top.grid_columnconfigure(0, weight=1)
    app.pdf_tool_progress = kit.progress(top)
    app.pdf_tool_progress.grid(row=0, column=0, sticky="ew", padx=(0, kit.pad("md")))
    kit.button(top, "Clear log", app._clear_pdf_tool_log, variant="ghost").grid(row=0, column=1, padx=(0, kit.pad("sm")))
    app.pdf_tool_start_button = kit.button(top, "Run PDF tool", app._start_pdf_tool, variant="primary")
    app.pdf_tool_start_button.grid(row=0, column=2)

    app.pdf_tool_log_text = kit.textbox(body, height=240)
    app.pdf_tool_log_text.grid(row=1, column=0, sticky="nsew")
    app.pdf_tool_log_text.configure(state="disabled")
