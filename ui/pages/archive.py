"""Archive tools page: zip, inspect, extract, and batch-compress.

Self-contained: this module owns its widgets *and* its handlers, so nothing has
to be threaded back through app.py. Long operations run on a worker thread and
report back through the app's Tk event loop, so the window never freezes.
"""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import archive_core
from archive_core import ArchiveSummary, format_size
from converter_core import ConversionError
from ui.dnd import DropZone, register_drop_target

COMPRESSION_LABELS = {
    "store": "Store (no compression, fastest)",
    "fast": "Fast",
    "normal": "Normal (recommended)",
    "maximum": "Maximum (smallest, slowest)",
}
_LABEL_TO_KEY = {label: key for key, label in COMPRESSION_LABELS.items()}


class ArchivePage:
    def __init__(self, app) -> None:
        self.app = app
        self.kit = app.kit
        self.sources: list[Path] = []
        self.summary: ArchiveSummary | None = None
        self.busy = False

        state = app.state_store
        self.output_dir_var = tk.StringVar(value=str(state.get("archive_output_dir", "") or ""))
        self.compression_var = tk.StringVar(
            value=COMPRESSION_LABELS.get(str(state.get("archive_compression", "normal")), COMPRESSION_LABELS["normal"])
        )
        self.subfolder_var = tk.BooleanVar(value=bool(state.get("archive_create_subfolder", True)))
        self.archive_path_var = tk.StringVar(value=str(state.get("archive_last_archive", "") or ""))
        self.status_var = tk.StringVar(value="Add files or folders to build an archive.")

        self._build()

    # -- construction -------------------------------------------------------

    def _build(self) -> None:
        kit = self.kit
        page = self.app._create_page_frame("archive", scrollable=True)
        page.grid_columnconfigure(0, weight=1)

        kit.page_header(
            page,
            "Archive",
            "Bundle files into a .zip, look inside one before you unpack it, or compress a batch of folders at once.",
        ).grid(row=0, column=0, sticky="ew", pady=(0, kit.pad("lg")))

        self._build_create_card(page).grid(row=1, column=0, sticky="ew", pady=(0, kit.pad("lg")))
        self._build_extract_card(page).grid(row=2, column=0, sticky="ew", pady=(0, kit.pad("lg")))
        self._build_log_card(page).grid(row=3, column=0, sticky="nsew")
        page.grid_rowconfigure(3, weight=1)

    def _build_create_card(self, parent):
        kit = self.kit
        card = kit.card(
            parent,
            title="Create an archive",
            subtitle="Files keep their names. A folder keeps its internal structure "
            "under a top-level folder of the same name.",
        )
        body = card.body
        body.grid_columnconfigure(0, weight=1)

        self.drop_zone = DropZone(
            body,
            kit,
            on_files=self.add_sources,
            on_browse=self.browse_files,
            title="Drop files or folders here",
            browse_text="Add files…",
        )
        self.drop_zone.grid(row=0, column=0, sticky="ew")

        list_holder = kit.frame(body, variant="transparent")
        list_holder.grid(row=1, column=0, sticky="nsew", pady=(kit.pad("md"), 0))
        list_holder.grid_columnconfigure(0, weight=1)

        self.source_list = kit.listbox(list_holder, selectmode=tk.EXTENDED, height=6)
        self.source_list.grid(row=0, column=0, sticky="nsew")
        register_drop_target(self.source_list, self.add_sources)

        side = kit.frame(list_holder, variant="transparent")
        side.grid(row=0, column=1, sticky="ns", padx=(kit.pad("md"), 0))
        for index, (text, command, variant) in enumerate(
            [
                ("Add files", self.browse_files, "secondary"),
                ("Add folder", self.browse_folder, "secondary"),
                ("Remove", self.remove_selected, "secondary"),
                ("Clear", self.clear_sources, "ghost"),
            ]
        ):
            kit.button(side, text, command, variant=variant, width=130).grid(
                row=index, column=0, sticky="ew", pady=(0 if index == 0 else kit.pad("sm"), 0)
            )

        options = kit.frame(body, variant="transparent")
        options.grid(row=2, column=0, sticky="ew", pady=(kit.pad("md"), 0))
        options.grid_columnconfigure(1, weight=1)

        row = kit.field_row(
            options,
            0,
            "Save archive to",
            kit.entry(options, self.output_dir_var, placeholder="Defaults to the first item's folder"),
            trailing=kit.button(options, "Browse…", self.browse_output_dir, variant="secondary", width=100),
        )
        row = kit.field_row(
            options,
            row,
            "Compression",
            kit.combo(options, list(COMPRESSION_LABELS.values()), self.compression_var),
        )

        actions = kit.frame(body, variant="transparent")
        actions.grid(row=3, column=0, sticky="ew", pady=(kit.pad("md"), 0))
        self.create_button = kit.button(actions, "Create archive", self.create_archive, variant="primary")
        self.create_button.grid(row=0, column=0)
        self.batch_button = kit.button(
            actions,
            "Batch-compress each folder",
            self.batch_compress,
            variant="secondary",
            tooltip="Every selected folder becomes its own .zip file.",
        )
        self.batch_button.grid(row=0, column=1, padx=(kit.pad("sm"), 0))
        return card

    def _build_extract_card(self, parent):
        kit = self.kit
        card = kit.card(parent, title="Open an archive", subtitle="Check what is inside before you extract it.")
        body = card.body
        body.grid_columnconfigure(0, weight=1)

        picker = kit.frame(body, variant="transparent")
        picker.grid(row=0, column=0, sticky="ew")
        picker.grid_columnconfigure(1, weight=1)
        kit.field_row(
            picker,
            0,
            "Archive",
            kit.entry(picker, self.archive_path_var, placeholder="Choose a .zip file"),
            trailing=kit.button(picker, "Browse…", self.browse_archive, variant="secondary", width=100),
        )

        self.contents_tree = kit.treeview(body, columns=("name", "size", "packed", "modified"), show="headings", height=8)
        for column, title, width, anchor in (
            ("name", "Name", 420, "w"),
            ("size", "Size", 90, "e"),
            ("packed", "Packed", 90, "e"),
            ("modified", "Modified", 140, "w"),
        ):
            self.contents_tree.heading(column, text=title)
            self.contents_tree.column(column, width=width, anchor=anchor)
        self.contents_tree.grid(row=1, column=0, sticky="nsew", pady=(kit.pad("sm"), 0))
        register_drop_target(self.contents_tree, self._drop_archive)

        self.summary_label = kit.label(body, "No archive loaded.", variant="small")
        self.summary_label.grid(row=2, column=0, sticky="w", pady=(kit.pad("sm"), 0))

        actions = kit.frame(body, variant="transparent")
        actions.grid(row=3, column=0, sticky="ew", pady=(kit.pad("md"), 0))
        kit.button(actions, "List contents", self.inspect_archive, variant="secondary").grid(row=0, column=0)
        self.extract_button = kit.button(actions, "Extract…", self.extract_archive, variant="primary")
        self.extract_button.grid(row=0, column=1, padx=(kit.pad("sm"), 0))
        kit.check(actions, "Extract into a folder named after the archive", self.subfolder_var).grid(
            row=0, column=2, padx=(kit.pad("lg"), 0)
        )
        return card

    def _build_log_card(self, parent):
        kit = self.kit
        card = kit.card(parent, title="Progress")
        body = card.body
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(2, weight=1)

        self.progress = kit.progress(body)
        self.progress.grid(row=0, column=0, sticky="ew")
        kit.label(body, "", variant="small").grid(row=1, column=0, sticky="w")
        self.status_label = kit.label(body, "", variant="small", textvariable=self.status_var)
        self.status_label.grid(row=1, column=0, sticky="w", pady=(kit.pad("sm"), 0))

        self.log = kit.textbox(body, height=140)
        self.log.grid(row=2, column=0, sticky="nsew", pady=(kit.pad("sm"), 0))
        self.log.configure(state="disabled")
        return card

    # -- state --------------------------------------------------------------

    def persist(self) -> None:
        self.app.state_store.update(
                archive_output_dir=self.output_dir_var.get().strip(),
                archive_compression=_LABEL_TO_KEY.get(self.compression_var.get(), "normal"),
                archive_create_subfolder=bool(self.subfolder_var.get()),
                archive_last_archive=self.archive_path_var.get().strip(),
        )

    def _refresh_source_list(self) -> None:
        self.source_list.delete(0, tk.END)
        for item in self.sources:
            marker = "[folder] " if item.is_dir() else ""
            self.source_list.insert(tk.END, f"{marker}{item}")
        count = len(self.sources)
        self.status_var.set(
            "Add files or folders to build an archive." if not count else f"{count} item(s) ready to compress."
        )

    def add_sources(self, paths: list[Path]) -> None:
        existing = {str(item) for item in self.sources}
        added = 0
        for path in paths:
            resolved = Path(path).expanduser()
            if resolved.exists() and str(resolved) not in existing:
                self.sources.append(resolved)
                existing.add(str(resolved))
                added += 1
        self._refresh_source_list()
        if added:
            self.status_var.set(f"Added {added} item(s).")

    def browse_files(self) -> None:
        selected = filedialog.askopenfilenames(title="Add files to the archive")
        if selected:
            self.add_sources([Path(item) for item in selected])

    def browse_folder(self) -> None:
        folder = filedialog.askdirectory(title="Add a folder to the archive")
        if folder:
            self.add_sources([Path(folder)])

    def browse_output_dir(self) -> None:
        folder = filedialog.askdirectory(title="Where should the archive be saved?")
        if folder:
            self.output_dir_var.set(folder)
            self.persist()

    def browse_archive(self) -> None:
        selected = filedialog.askopenfilename(
            title="Open a zip archive", filetypes=[("Zip archives", "*.zip"), ("All files", "*.*")]
        )
        if selected:
            self.archive_path_var.set(selected)
            self.persist()
            self.inspect_archive()

    def _drop_archive(self, paths: list[Path]) -> None:
        for path in paths:
            if path.suffix.lower() == ".zip":
                self.archive_path_var.set(str(path))
                self.inspect_archive()
                return
        self.status_var.set("Drop a .zip file here to inspect it.")

    def remove_selected(self) -> None:
        for index in sorted(self.source_list.curselection(), reverse=True):
            del self.sources[index]
        self._refresh_source_list()

    def clear_sources(self) -> None:
        self.sources.clear()
        self._refresh_source_list()

    # -- logging / worker plumbing ------------------------------------------

    def _append_log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", f"{message}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def log_line(self, message: str) -> None:
        """Thread-safe log write."""
        self.app.after(0, self._append_log, message)

    def report_progress(self, current: int, total: int) -> None:
        value = 0.0 if not total else min(1.0, current / total)
        self.app.after(0, self.progress.set, value)

    def _set_busy(self, busy: bool) -> None:
        self.busy = busy
        for button in (self.create_button, self.batch_button, self.extract_button):
            try:
                button.configure(state="disabled" if busy else "normal")
            except Exception:
                pass

    def _run_worker(self, label: str, work, on_success) -> None:
        """Run ``work`` off the UI thread; ``on_success(result)`` runs back on it."""
        if self.busy:
            messagebox.showinfo("Archive", "Another archive operation is still running.")
            return
        self._set_busy(True)
        self.progress.set(0)
        self.status_var.set(f"{label}…")

        def runner() -> None:
            try:
                result = work()
            except ConversionError as exc:
                self.app.after(0, self._fail, str(exc))
            except Exception as exc:  # noqa: BLE001 - surfaced to the user, not swallowed
                self.app.after(0, self._fail, f"Unexpected problem: {exc}")
            else:
                self.app.after(0, self._succeed, result, on_success)

        threading.Thread(target=runner, daemon=True).start()

    def _fail(self, message: str) -> None:
        self._set_busy(False)
        self.progress.set(0)
        self.status_var.set("Did not finish — see the message.")
        self._append_log(f"ERROR: {message}")
        messagebox.showerror("Archive", message)

    def _succeed(self, result, on_success) -> None:
        self._set_busy(False)
        self.progress.set(1.0)
        on_success(result)

    # -- operations ---------------------------------------------------------

    def _resolve_output_dir(self) -> Path:
        raw = self.output_dir_var.get().strip()
        if raw:
            return Path(raw).expanduser()
        if self.sources:
            return self.sources[0].parent
        return Path.cwd()

    def create_archive(self) -> None:
        if not self.sources:
            messagebox.showinfo("Archive", "Add at least one file or folder first.")
            return
        self.persist()
        default_name = archive_core.suggest_archive_name(self.sources)
        target = filedialog.asksaveasfilename(
            title="Save archive as",
            defaultextension=".zip",
            initialfile=default_name,
            initialdir=str(self._resolve_output_dir()),
            filetypes=[("Zip archives", "*.zip")],
        )
        if not target:
            return
        compression = _LABEL_TO_KEY.get(self.compression_var.get(), "normal")
        sources = list(self.sources)

        self._run_worker(
            "Compressing",
            lambda: archive_core.create_archive(
                sources,
                target,
                compression=compression,
                overwrite=True,
                log=self.log_line,
                progress=self.report_progress,
            ),
            self._after_create,
        )

    def _after_create(self, path: Path) -> None:
        self.status_var.set(f"Created {path.name} ({format_size(path.stat().st_size)}).")
        self.archive_path_var.set(str(path))
        self.persist()
        self.app._track_success_outputs([path], output_dir=path.parent, label="Archive")
        self.inspect_archive()

    def batch_compress(self) -> None:
        folders = [item for item in self.sources if item.is_dir()]
        if not folders:
            messagebox.showinfo(
                "Batch compress",
                "Add one or more folders first — batch compress makes one .zip per folder.",
            )
            return
        output_dir = filedialog.askdirectory(title="Where should the archives be saved?", initialdir=str(self._resolve_output_dir()))
        if not output_dir:
            return
        self.output_dir_var.set(output_dir)
        self.persist()
        compression = _LABEL_TO_KEY.get(self.compression_var.get(), "normal")

        self._run_worker(
            f"Compressing {len(folders)} folder(s)",
            lambda: archive_core.batch_compress(
                folders,
                output_dir,
                compression=compression,
                overwrite=True,
                log=self.log_line,
                progress=self.report_progress,
            ),
            self._after_batch,
        )

    def _after_batch(self, outputs: list[Path]) -> None:
        self.status_var.set(f"Created {len(outputs)} archive(s).")
        if outputs:
            self.app._track_success_outputs(outputs, output_dir=outputs[0].parent, label="Batch compress")

    def inspect_archive(self) -> None:
        raw = self.archive_path_var.get().strip()
        if not raw:
            messagebox.showinfo("Archive", "Choose a .zip file to look inside.")
            return
        try:
            summary = archive_core.read_archive(raw)
        except ConversionError as exc:
            self.summary = None
            self._render_contents(None)
            messagebox.showerror("Archive", str(exc))
            return

        self.summary = summary
        self._render_contents(summary)
        self.persist()

    def _render_contents(self, summary: ArchiveSummary | None) -> None:
        for item in self.contents_tree.get_children():
            self.contents_tree.delete(item)
        if summary is None:
            self.summary_label.configure(text="No archive loaded.")
            return
        for entry in summary.entries:
            self.contents_tree.insert(
                "",
                "end",
                values=(
                    entry.name,
                    "" if entry.is_dir else format_size(entry.size),
                    "" if entry.is_dir else format_size(entry.compressed_size),
                    entry.modified,
                ),
            )
        self.summary_label.configure(
            text=(
                f"{summary.file_count} file(s) · {format_size(summary.total_size)} uncompressed · "
                f"{format_size(summary.total_compressed)} packed · {summary.ratio * 100:.0f}% saved"
            )
        )
        self.status_var.set(f"Loaded {summary.path.name}.")

    def extract_archive(self) -> None:
        raw = self.archive_path_var.get().strip()
        if not raw:
            messagebox.showinfo("Archive", "Choose a .zip file to extract.")
            return
        destination = filedialog.askdirectory(title="Extract into which folder?", initialdir=str(self._resolve_output_dir()))
        if not destination:
            return
        create_subfolder = bool(self.subfolder_var.get())
        self.persist()

        self._run_worker(
            "Extracting",
            lambda: archive_core.extract_archive(
                raw,
                destination,
                create_subfolder=create_subfolder,
                log=self.log_line,
                progress=self.report_progress,
            ),
            self._after_extract,
        )

    def _after_extract(self, folder: Path) -> None:
        self.status_var.set(f"Extracted into {folder}.")
        self.app._track_success_outputs([folder], output_dir=folder, label="Extract archive")


def build(app) -> ArchivePage:
    page = ArchivePage(app)
    app.archive_page = page
    return page
