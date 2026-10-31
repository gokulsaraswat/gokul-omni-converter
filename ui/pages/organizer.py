"""Organizer page — hosts the visual page-reordering panel.

PageOrganizerPanel is a ttk panel with its own palette handling
(``apply_theme``), so this page just gives it a full-bleed CTk container.
"""

from __future__ import annotations

from page_organizer import PageOrganizerPanel


def build(app) -> None:
    page = app._create_page_frame("organizer", scrollable=False)
    page.grid_columnconfigure(0, weight=1)
    page.grid_rowconfigure(0, weight=1)

    app.organizer_panel = PageOrganizerPanel(
        page,
        palette=app._current_palette(),
        set_status=app.status_var.set,
        on_recent_job=app._record_external_job,
        on_loaded_pdf=app._remember_organizer_pdf,
        kit=app.kit,
    )
    app.organizer_panel.grid(row=0, column=0, sticky="nsew")
