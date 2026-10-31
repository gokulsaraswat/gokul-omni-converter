"""Batch rename and folder organiser page.

Everything here is preview-first: the user always sees the exact old -> new list,
with problems called out per row, before anything moves. The last applied batch
can always be undone from the same screen.
"""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import rename_organizer_core as roc
from converter_core import ConversionError
from rename_organizer_core import OperationPlan
from ui.dnd import DropZone, register_drop_target

RENAME_MODE_LABELS = {
    "pattern": "Pattern with tokens",
    "find_replace": "Find and replace",
    "regex": "Regular expression",
    "case": "Change case",
}
_RENAME_LABEL_TO_KEY = {label: key for key, label in RENAME_MODE_LABELS.items()}

CASE_LABELS = {
    "lower": "lowercase",
    "upper": "UPPERCASE",
    "title": "Title Case",
    "kebab": "kebab-case",
    "snake": "snake_case",
}
_CASE_LABEL_TO_KEY = {label: key for key, label in CASE_LABELS.items()}

SORT_LABELS = {
    "name": "Name",
    "modified": "Date modified",
    "created": "Date created",
    "size": "Size",
    "extension": "Extension",
}
_SORT_LABEL_TO_KEY = {label: key for key, label in SORT_LABELS.items()}

ORGANIZE_LABELS = {
    "type": "File type (Documents, Images, …)",
    "extension": "Extension (PDF, JPG, …)",
    "date_modified": "Month modified (2026-08)",
    "date_created": "Month created (2026-08)",
}
_ORGANIZE_LABEL_TO_KEY = {label: key for key, label in ORGANIZE_LABELS.items()}


class RenameOrganizePage:
    def __init__(self, app) -> None:
        self.app = app
        self.kit = app.kit
        self.files: list[Path] = []
        self.plan: OperationPlan | None = None
        self.busy = False

        state = app.state_store
        self.mode_var = tk.StringVar(
            value=RENAME_MODE_LABELS.get(str(state.get("rename_mode", "pattern")), RENAME_MODE_LABELS["pattern"])
        )
        self.pattern_var = tk.StringVar(value=str(state.get("rename_pattern", "{name}")))
        self.find_var = tk.StringVar(value=str(state.get("rename_find", "")))
        self.replace_var = tk.StringVar(value=str(state.get("rename_replace", "")))
        self.regex_var = tk.StringVar(value=str(state.get("rename_regex", "")))
        self.regex_replacement_var = tk.StringVar(value=str(state.get("rename_regex_replacement", "")))
        self.case_var = tk.StringVar(
            value=CASE_LABELS.get(str(state.get("rename_case_mode", "lower")), CASE_LABELS["lower"])
        )
        self.sort_var = tk.StringVar(
            value=SORT_LABELS.get(str(state.get("rename_sort_mode", "name")), SORT_LABELS["name"])
        )
        self.start_number_var = tk.StringVar(value=str(state.get("rename_start_number", 1)))
        self.padding_var = tk.StringVar(value=str(state.get("rename_number_padding", 4)))
        self.keep_extension_var = tk.BooleanVar(value=bool(state.get("rename_keep_extension", True)))

        self.organize_folder_var = tk.StringVar(value=str(state.get("organize_folder", "") or ""))
        self.organize_mode_var = tk.StringVar(
            value=ORGANIZE_LABELS.get(str(state.get("organize_mode", "type")), ORGANIZE_LABELS["type"])
        )
        self.organize_recursive_var = tk.BooleanVar(value=bool(state.get("organize_recursive", False)))

        self.status_var = tk.StringVar(value="Add files, then preview before anything is renamed.")
        self.undo_var = tk.StringVar(value="")

        self._build()
        self.refresh_undo_label()

    # -- construction -------------------------------------------------------

    def _build(self) -> None:
        kit = self.kit
        page = self.app._create_page_frame("rename_organize", scrollable=True)
        page.grid_columnconfigure(0, weight=1)

        kit.page_header(
            page,
            "Rename & Organize",
            "Rename a batch of files by pattern, or sort a folder into subfolders. "
            "Nothing moves until you approve the preview, and the last batch can always be undone.",
        ).grid(row=0, column=0, sticky="ew", pady=(0, kit.pad("lg")))

        self.tabs = kit.segmented(page, ["Batch rename", "Organize a folder"], command=self._switch_tab)
        self.tabs.grid(row=1, column=0, sticky="w", pady=(0, kit.pad("md")))
        self.tabs.set("Batch rename")

        self.rename_card = self._build_rename_card(page)
        self.rename_card.grid(row=2, column=0, sticky="ew")
        self.organize_card = self._build_organize_card(page)

        self._build_preview_card(page).grid(row=3, column=0, sticky="nsew", pady=(kit.pad("lg"), 0))
        page.grid_rowconfigure(3, weight=1)

    def _switch_tab(self, value: str) -> None:
        if value.startswith("Batch"):
            self.organize_card.grid_forget()
            self.rename_card.grid(row=2, column=0, sticky="ew")
        else:
            self.rename_card.grid_forget()
            self.organize_card.grid(row=2, column=0, sticky="ew")
        self._clear_preview()

    def _build_rename_card(self, parent):
        kit = self.kit
        card = kit.card(parent, title="Files to rename")
        body = card.body
        body.grid_columnconfigure(0, weight=1)

        self.drop_zone = DropZone(
            body,
            kit,
            on_files=self.add_files,
            on_browse=self.browse_files,
            title="Drop the files you want to rename",
            browse_text="Choose files…",
        )
        self.drop_zone.grid(row=0, column=0, sticky="ew")

        list_holder = kit.frame(body, variant="transparent")
        list_holder.grid(row=1, column=0, sticky="ew", pady=(kit.pad("md"), 0))
        list_holder.grid_columnconfigure(0, weight=1)
        self.file_list = kit.listbox(list_holder, selectmode=tk.EXTENDED, height=5)
        self.file_list.grid(row=0, column=0, sticky="nsew")
        register_drop_target(self.file_list, self.add_files)

        side = kit.frame(list_holder, variant="transparent")
        side.grid(row=0, column=1, sticky="ns", padx=(kit.pad("md"), 0))
        for index, (text, command, variant) in enumerate(
            [
                ("Add files", self.browse_files, "secondary"),
                ("Add folder", self.browse_folder, "secondary"),
                ("Remove", self.remove_selected, "secondary"),
                ("Clear", self.clear_files, "ghost"),
            ]
        ):
            kit.button(side, text, command, variant=variant, width=130).grid(
                row=index, column=0, sticky="ew", pady=(0 if index == 0 else kit.pad("sm"), 0)
            )

        rules = kit.frame(body, variant="transparent")
        rules.grid(row=2, column=0, sticky="ew", pady=(kit.pad("md"), 0))
        rules.grid_columnconfigure(1, weight=1)

        row = kit.field_row(rules, 0, "Rule", kit.combo(rules, list(RENAME_MODE_LABELS.values()), self.mode_var, command=lambda _v: self._on_rule_changed()))
        row = kit.field_row(
            rules,
            row,
            "Pattern",
            kit.entry(rules, self.pattern_var, placeholder="{name} or invoice_{n}"),
            hint="",  # token reference moved to the ⓘ on the Pattern label
        )
        row = kit.field_row(rules, row, "Find", kit.entry(rules, self.find_var, placeholder="Text to look for"))
        row = kit.field_row(rules, row, "Replace with", kit.entry(rules, self.replace_var, placeholder="Leave blank to delete it"))
        row = kit.field_row(rules, row, "Expression", kit.entry(rules, self.regex_var, placeholder=r"^IMG_(\d+)"))
        row = kit.field_row(rules, row, "Replacement", kit.entry(rules, self.regex_replacement_var, placeholder=r"photo_\1"))
        row = kit.field_row(rules, row, "Case", kit.combo(rules, list(CASE_LABELS.values()), self.case_var))

        numbering = kit.frame(rules, variant="transparent")
        numbering.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(kit.pad("sm"), 0))
        kit.label(numbering, "Order by", variant="body").grid(row=0, column=0, padx=(0, kit.pad("sm")))
        kit.combo(numbering, list(SORT_LABELS.values()), self.sort_var, width=140).grid(row=0, column=1)
        kit.label(numbering, "Start at", variant="body").grid(row=0, column=2, padx=(kit.pad("lg"), kit.pad("sm")))
        kit.entry(numbering, self.start_number_var, width=70).grid(row=0, column=3)
        kit.label(numbering, "Digits", variant="body").grid(row=0, column=4, padx=(kit.pad("lg"), kit.pad("sm")))
        kit.entry(numbering, self.padding_var, width=70).grid(row=0, column=5)
        kit.check(numbering, "Keep original extension", self.keep_extension_var).grid(
            row=0, column=6, padx=(kit.pad("lg"), 0)
        )

        # Rows whose visibility depends on the chosen rule.
        self._rule_rows = {
            "pattern": (1, 2),
            "find_replace": (3, 4),
            "regex": (5, 6),
            "case": (7,),
        }
        self._rules_frame = rules
        self._on_rule_changed()
        return card

    def _on_rule_changed(self) -> None:
        """Only show the inputs the selected rule actually uses."""
        active = _RENAME_LABEL_TO_KEY.get(self.mode_var.get(), "pattern")
        visible = set(self._rule_rows.get(active, ()))
        for key, rows in self._rule_rows.items():
            for grid_row in rows:
                for child in self._rules_frame.grid_slaves(row=grid_row):
                    if key == active or grid_row in visible:
                        child.grid()
                    else:
                        child.grid_remove()

    def _build_organize_card(self, parent):
        kit = self.kit
        card = kit.card(
            parent,
            title="Organize a folder",
            subtitle="Files are moved into subfolders inside the same folder. Files already in a subfolder are left alone.",
        )
        body = card.body
        body.grid_columnconfigure(1, weight=1)

        row = kit.field_row(
            body,
            0,
            "Folder",
            kit.entry(body, self.organize_folder_var, placeholder="Choose the folder to tidy up"),
            trailing=kit.button(body, "Browse…", self.browse_organize_folder, variant="secondary", width=100),
        )
        row = kit.field_row(body, row, "Sort into", kit.combo(body, list(ORGANIZE_LABELS.values()), self.organize_mode_var))
        kit.check(body, "Include files in existing subfolders", self.organize_recursive_var).grid(
            row=row, column=1, sticky="w", pady=kit.pad("sm")
        )
        return card

    def _build_preview_card(self, parent):
        kit = self.kit
        card = kit.card(parent, title="Preview", subtitle="Check this list before applying. Rows with a problem are skipped.")
        body = card.body
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(1, weight=1)

        actions = kit.frame(body, variant="transparent")
        actions.grid(row=0, column=0, sticky="ew")
        kit.button(actions, "Preview changes", self.preview, variant="secondary").grid(row=0, column=0)
        self.apply_button = kit.button(actions, "Apply", self.apply, variant="primary")
        self.apply_button.grid(row=0, column=1, padx=(kit.pad("sm"), 0))
        self.apply_button.configure(state="disabled")
        self.undo_button = kit.button(actions, "Undo last operation", self.undo, variant="danger")
        self.undo_button.grid(row=0, column=2, padx=(kit.pad("lg"), 0))

        self.preview_tree = kit.treeview(body, columns=("old", "new", "note"), show="headings", height=10)
        for column, title, width in (("old", "Current name", 300), ("new", "New name", 300), ("note", "Note", 240)):
            self.preview_tree.heading(column, text=title)
            self.preview_tree.column(column, width=width, anchor="w")
        self.preview_tree.grid(row=1, column=0, sticky="nsew", pady=(kit.pad("md"), 0))

        self.empty_state = kit.empty_state(
            body,
            title="Nothing previewed yet",
            body="Pick your files and a rule above, then choose Preview changes. "
            "You will see every old and new name here before anything is touched.",
        )
        self.empty_state.grid(row=1, column=0, sticky="nsew", pady=(kit.pad("md"), 0))

        kit.label(body, "", variant="small", textvariable=self.status_var).grid(
            row=2, column=0, sticky="w", pady=(kit.pad("sm"), 0)
        )
        kit.label(body, "", variant="tiny", textvariable=self.undo_var).grid(row=3, column=0, sticky="w")
        return card

    # -- state --------------------------------------------------------------

    def persist(self) -> None:
        self.app.state_store.update(
                rename_mode=_RENAME_LABEL_TO_KEY.get(self.mode_var.get(), "pattern"),
                rename_pattern=self.pattern_var.get(),
                rename_find=self.find_var.get(),
                rename_replace=self.replace_var.get(),
                rename_regex=self.regex_var.get(),
                rename_regex_replacement=self.regex_replacement_var.get(),
                rename_case_mode=_CASE_LABEL_TO_KEY.get(self.case_var.get(), "lower"),
                rename_sort_mode=_SORT_LABEL_TO_KEY.get(self.sort_var.get(), "name"),
                rename_start_number=self._int(self.start_number_var.get(), 1),
                rename_number_padding=self._int(self.padding_var.get(), 4),
                rename_keep_extension=bool(self.keep_extension_var.get()),
                organize_folder=self.organize_folder_var.get().strip(),
                organize_mode=_ORGANIZE_LABEL_TO_KEY.get(self.organize_mode_var.get(), "type"),
                organize_recursive=bool(self.organize_recursive_var.get()),
        )

    @staticmethod
    def _int(value: str, fallback: int) -> int:
        try:
            return int(str(value).strip())
        except (TypeError, ValueError):
            return fallback

    def _refresh_file_list(self) -> None:
        self.file_list.delete(0, tk.END)
        for item in self.files:
            self.file_list.insert(tk.END, str(item))
        self.status_var.set(
            "Add files, then preview before anything is renamed."
            if not self.files
            else f"{len(self.files)} file(s) selected."
        )

    def add_files(self, paths: list[Path]) -> None:
        existing = {str(item) for item in self.files}
        for path in paths:
            resolved = Path(path).expanduser()
            if resolved.is_dir():
                for child in sorted(resolved.glob("*")):
                    if child.is_file() and str(child) not in existing:
                        self.files.append(child)
                        existing.add(str(child))
            elif resolved.is_file() and str(resolved) not in existing:
                self.files.append(resolved)
                existing.add(str(resolved))
        self._refresh_file_list()
        self._clear_preview()

    def browse_files(self) -> None:
        selected = filedialog.askopenfilenames(title="Choose files to rename")
        if selected:
            self.add_files([Path(item) for item in selected])

    def browse_folder(self) -> None:
        folder = filedialog.askdirectory(title="Add every file in a folder")
        if folder:
            self.add_files([Path(folder)])

    def browse_organize_folder(self) -> None:
        folder = filedialog.askdirectory(title="Choose the folder to organise")
        if folder:
            self.organize_folder_var.set(folder)
            self.persist()
            self._clear_preview()

    def remove_selected(self) -> None:
        for index in sorted(self.file_list.curselection(), reverse=True):
            del self.files[index]
        self._refresh_file_list()
        self._clear_preview()

    def clear_files(self) -> None:
        self.files.clear()
        self._refresh_file_list()
        self._clear_preview()

    # -- preview / apply ----------------------------------------------------

    def _is_rename_tab(self) -> bool:
        return str(self.tabs.get()).startswith("Batch")

    def _clear_preview(self) -> None:
        self.plan = None
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        self.apply_button.configure(state="disabled")
        self.preview_tree.grid_remove()
        self.empty_state.grid()

    def _build_plan(self) -> OperationPlan:
        if self._is_rename_tab():
            return roc.plan_rename(
                self.files,
                mode=_RENAME_LABEL_TO_KEY.get(self.mode_var.get(), "pattern"),
                pattern=self.pattern_var.get(),
                find_text=self.find_var.get(),
                replace_text=self.replace_var.get(),
                regex_pattern=self.regex_var.get(),
                regex_replacement=self.regex_replacement_var.get(),
                case_mode=_CASE_LABEL_TO_KEY.get(self.case_var.get(), "lower"),
                keep_extension=bool(self.keep_extension_var.get()),
                sort_mode=_SORT_LABEL_TO_KEY.get(self.sort_var.get(), "name"),
                start_number=self._int(self.start_number_var.get(), 1),
                number_padding=self._int(self.padding_var.get(), 4),
            )
        return roc.plan_organize(
            self.organize_folder_var.get().strip(),
            mode=_ORGANIZE_LABEL_TO_KEY.get(self.organize_mode_var.get(), "type"),
            recursive=bool(self.organize_recursive_var.get()),
        )

    def preview(self) -> None:
        self.persist()
        try:
            plan = self._build_plan()
        except ConversionError as exc:
            self._clear_preview()
            messagebox.showerror("Preview", str(exc))
            return

        self.plan = plan
        self.empty_state.grid_remove()
        self.preview_tree.grid()
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)

        organize = plan.kind == "organize"
        for move in plan.moves:
            if move.problem:
                note, tag = move.problem, "error"
            elif not move.changed:
                note, tag = "No change", "warn"
            else:
                note, tag = "Ready", "success"
            new_label = str(move.target.relative_to(plan.root)) if organize and move.changed else move.target.name
            self.preview_tree.insert("", "end", values=(move.source.name, new_label, note), tags=(tag,))

        ready = len(plan.applicable)
        problems = len(plan.problems)
        self.apply_button.configure(state="normal" if ready else "disabled")
        summary = f"{ready} file(s) will change"
        if problems:
            summary += f", {problems} skipped — see the Note column"
        if not ready:
            summary = (
                "Nothing would change with these settings. Adjust the rule above and preview again."
                if not problems
                else f"No file can be changed — {problems} row(s) have a problem listed in the Note column."
            )
        self.status_var.set(summary)

    def apply(self) -> None:
        if self.plan is None:
            messagebox.showinfo("Apply", "Preview the changes first.")
            return
        plan = self.plan
        count = len(plan.applicable)
        if not messagebox.askyesno(
            "Apply changes",
            f"{plan.description} will move or rename {count} file(s).\n\n"
            "You can reverse this with Undo last operation.\n\nGo ahead?",
        ):
            return
        if self.busy:
            return
        self.busy = True
        self.apply_button.configure(state="disabled")
        state_dir = self.app.state_store.path.parent

        def runner() -> None:
            try:
                applied, _manifest = roc.apply_plan(plan, state_dir)
            except ConversionError as exc:
                self.app.after(0, self._finish_apply, 0, str(exc))
            except Exception as exc:  # noqa: BLE001
                self.app.after(0, self._finish_apply, 0, f"Unexpected problem: {exc}")
            else:
                self.app.after(0, self._finish_apply, len(applied), "")

        threading.Thread(target=runner, daemon=True).start()

    def _finish_apply(self, count: int, error: str) -> None:
        self.busy = False
        self.refresh_undo_label()
        if error:
            self.status_var.set("Did not finish — see the message.")
            messagebox.showerror("Apply", error)
        else:
            self.status_var.set(f"Applied to {count} file(s). Use Undo last operation to reverse it.")
        # Paths on disk have moved, so the old selection and preview are stale.
        self.files.clear()
        self._refresh_file_list()
        self._clear_preview()

    def undo(self) -> None:
        state_dir = self.app.state_store.path.parent
        latest = roc.latest_manifest(state_dir)
        if latest is None:
            messagebox.showinfo("Undo", "There is no recent rename or organise operation to undo.")
            return
        if not messagebox.askyesno("Undo", f"Reverse this operation?\n\n{roc.describe_manifest(latest)}"):
            return
        try:
            restored = roc.undo_last(state_dir)
        except ConversionError as exc:
            self.refresh_undo_label()
            messagebox.showerror("Undo", str(exc))
            return
        self.status_var.set(f"Restored {restored} file(s).")
        self.refresh_undo_label()

    def refresh_undo_label(self) -> None:
        latest = roc.latest_manifest(self.app.state_store.path.parent)
        if latest is None:
            self.undo_var.set("No operation available to undo.")
            self.undo_button.configure(state="disabled")
        else:
            self.undo_var.set(f"Undo available: {roc.describe_manifest(latest)}")
            self.undo_button.configure(state="normal")


def build(app) -> RenameOrganizePage:
    page = RenameOrganizePage(app)
    app.rename_organize_page = page
    return page
