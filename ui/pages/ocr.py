"""OCR page.

Written for someone who does not know what OCR, DPI, or a page-segmentation mode
is. The screen shows three plain-language actions; every piece of jargon is
tucked behind an ⓘ instead of printed on the page.
"""

from __future__ import annotations

from pathlib import Path

from ui.dnd import register_drop_target

HELP = {
    "page": "OCR reads text out of scans and photos.\n\n"
            "Pick an image or PDF, then choose what you want back: a PDF you can "
            "search inside, or a plain text file.",
    "input": "The scan, photo, or PDF you want to read text from.\n"
             "You can also drag a file onto this box.",
    "output": "Where the finished file gets saved.",
    "language": "Which language the text is written in.\n\n"
                "'eng' is English. Other languages need that language pack "
                "installed with Tesseract.",
    "dpi": "How finely the page is scanned before reading, in dots per inch.\n\n"
           "220 suits most documents. Raise it for small or faint print; "
           "lower it if OCR is taking too long.",
    "psm": "Page segmentation mode — how Tesseract expects the text to be laid out.\n\n"
           "6 means 'one uniform block of text' and works for most documents. "
           "Leave it alone unless results look scrambled.",
    "tesseract": "Tesseract is the free OCR engine that does the actual reading.\n\n"
                 "GO Omni Pro finds it automatically if it is installed. "
                 "Set the path by hand only if detection fails.",
    "image_pdf": "Best for a single scan or screenshot.\n\n"
                 "You get a PDF that looks identical but whose text can be "
                 "selected and searched.",
    "pdf_pdf": "Best for a PDF made of scanned pages.\n\n"
               "Each page is re-read and an invisible text layer is added, so "
               "Ctrl-F works across the whole document.",
    "text": "Gives you just the words, as a plain .txt file — handy for copying "
            "into another app.",
}


def build(app) -> None:
    import app as app_module

    kit = app.kit
    page = app._create_page_frame("ocr", scrollable=True)
    page.grid_columnconfigure(0, weight=1)
    page.grid_rowconfigure(2, weight=1)

    kit.page_header(page, "OCR", HELP["page"]).grid(
        row=0, column=0, sticky="ew", pady=(0, kit.pad("lg"))
    )

    content = kit.frame(page, variant="transparent")
    content.grid(row=1, column=0, sticky="nsew", pady=(0, kit.pad("lg")))
    content.grid_columnconfigure(0, weight=3)
    content.grid_columnconfigure(1, weight=2)
    content.grid_rowconfigure(0, weight=1)

    _build_source(app, kit, content, app_module)
    _build_settings(app, kit, content)
    _build_log(app, kit, page)


def _build_source(app, kit, content, app_module) -> None:
    card = kit.card(content, title="What do you want to read?")
    card.grid(row=0, column=0, sticky="nsew", padx=(0, kit.pad("md")))
    body = card.body
    body.grid_columnconfigure(1, weight=1)

    kit.labelled(body, "File", HELP["input"]).grid(row=0, column=0, sticky="w", padx=(0, kit.pad("sm")))
    input_entry = kit.entry(body, app.ocr_input_var, placeholder="Choose an image or PDF")
    input_entry.grid(row=0, column=1, sticky="ew")
    register_drop_target(input_entry, lambda paths: app.ocr_input_var.set(str(paths[0])))
    kit.button(body, "Browse", app._browse_ocr_input, variant="secondary", width=100).grid(
        row=0, column=2, padx=(kit.pad("sm"), 0)
    )

    kit.labelled(body, "Save to", HELP["output"]).grid(
        row=1, column=0, sticky="w", padx=(0, kit.pad("sm")), pady=(kit.pad("sm"), 0)
    )
    kit.entry(body, app.ocr_output_var).grid(row=1, column=1, sticky="ew", pady=(kit.pad("sm"), 0))
    output_actions = kit.button_bar(body, columns=2)
    output_actions.grid(row=1, column=2, sticky="e", padx=(kit.pad("sm"), 0), pady=(kit.pad("sm"), 0))
    output_actions.button("Browse", app._browse_ocr_output)
    output_actions.button(
        "Open",
        lambda: app_module.open_path(Path(app.ocr_output_var.get().strip() or Path.cwd())),
        variant="ghost",
    )

    kit.label(body, "Then pick one:", variant="subheading").grid(
        row=2, column=0, columnspan=3, sticky="w", pady=(kit.pad("xl"), kit.pad("sm"))
    )

    # Three plain-language actions, each with its own ⓘ.
    for index, (label, command, help_key, variant) in enumerate(
        [
            ("Make a searchable PDF from an image", app._start_ocr_image_pdf, "image_pdf", "primary"),
            ("Make a searchable PDF from a PDF", app._start_ocr_pdf_pdf, "pdf_pdf", "secondary"),
            ("Just give me the text", app._start_ocr_text, "text", "secondary"),
        ]
    ):
        row = kit.frame(body, variant="transparent")
        row.grid(row=3 + index, column=0, columnspan=3, sticky="ew", pady=(0, kit.pad("sm")))
        row.grid_columnconfigure(0, weight=1)
        button = kit.button(row, label, command, variant=variant, anchor="w")
        button.grid(row=0, column=0, sticky="ew")
        kit.info(row, HELP[help_key]).grid(row=0, column=1, padx=(kit.pad("sm"), 0))
        setattr(app, f"ocr_start_{['image', 'pdf', 'text'][index]}_button", button)


def _build_settings(app, kit, content) -> None:
    card = kit.card(content, title="Fine-tuning")
    card.grid(row=0, column=1, sticky="nsew")
    body = card.body
    body.grid_columnconfigure(1, weight=1)

    kit.label(
        body,
        "The defaults work for most documents — you can ignore this panel.",
        variant="small",
        wraplength=300,
        justify="left",
        anchor="w",
    ).grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, kit.pad("md")))

    row = 1
    for label, variable, help_key in (
        ("Language", app.ocr_language_var, "language"),
        ("Scan quality", app.ocr_dpi_var, "dpi"),
        ("Layout mode", app.ocr_psm_var, "psm"),
    ):
        kit.labelled(body, label, HELP[help_key]).grid(
            row=row, column=0, sticky="w", padx=(0, kit.pad("sm")), pady=(kit.pad("sm"), 0)
        )
        kit.entry(body, variable, width=110).grid(row=row, column=1, sticky="w", pady=(kit.pad("sm"), 0))
        row += 1

    kit.separator(body).grid(row=row, column=0, columnspan=3, sticky="ew", pady=kit.pad("lg"))
    row += 1

    kit.labelled(body, "OCR engine", HELP["tesseract"]).grid(
        row=row, column=0, sticky="w", padx=(0, kit.pad("sm"))
    )
    kit.entry(body, app.tesseract_path_var, placeholder="Found automatically").grid(
        row=row, column=1, columnspan=2, sticky="ew"
    )
    row += 1
    engine_actions = kit.button_bar(body, columns=2)
    engine_actions.grid(row=row, column=1, columnspan=2, sticky="w", pady=(kit.pad("sm"), 0))
    engine_actions.button("Locate…", app._browse_tesseract_path)
    engine_actions.button("Check it works", app._test_tesseract_path)
    row += 1

    for variable in (app.ocr_dependency_var, app.ocr_status_var):
        kit.label(
            body, "", variant="small", textvariable=variable,
            wraplength=300, justify="left", anchor="w",
        ).grid(row=row, column=0, columnspan=3, sticky="ew", pady=(kit.pad("md"), 0))
        row += 1


def _build_log(app, kit, page) -> None:
    card = kit.card(page, title="Progress")
    card.grid(row=2, column=0, sticky="nsew")
    body = card.body
    body.grid_rowconfigure(1, weight=1)

    top = kit.frame(body, variant="transparent")
    top.grid(row=0, column=0, sticky="ew", pady=(0, kit.pad("sm")))
    top.grid_columnconfigure(0, weight=1)
    app.ocr_progress = kit.progress(top)
    app.ocr_progress.grid(row=0, column=0, sticky="ew", padx=(0, kit.pad("md")))
    kit.button(top, "Clear", app._clear_ocr_log, variant="ghost").grid(row=0, column=1)

    app.ocr_log_text = kit.textbox(body, height=200)
    app.ocr_log_text.grid(row=1, column=0, sticky="nsew")
    app.ocr_log_text.configure(state="disabled")
