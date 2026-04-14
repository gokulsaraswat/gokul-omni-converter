from __future__ import annotations

import argparse
import queue
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from patch10_services import (
    OcrConfig,
    Patch10Error,
    SmtpConfig,
    build_eml_draft,
    compress_pdf,
    extract_text_with_ocr,
    image_to_searchable_pdf,
    open_mailto_draft,
    password_protect_pdf,
    pdf_to_searchable_pdf,
    redact_text,
    remove_pdf_password,
    send_email_smtp,
)

APP_TITLE = "Gokul Omni Convert Lite - Patch 10 Demo"


class Patch10DemoApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1240x860")
        self.minsize(1080, 760)
        self.worker_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.running = False

        self.ocr_input_var = tk.StringVar()
        self.ocr_output_var = tk.StringVar(value=str(Path.cwd() / "patch10_output" / "ocr"))
        self.ocr_lang_var = tk.StringVar(value="eng")
        self.ocr_dpi_var = tk.StringVar(value="220")
        self.ocr_psm_var = tk.StringVar(value="6")

        self.pdf_input_var = tk.StringVar()
        self.pdf_output_var = tk.StringVar(value=str(Path.cwd() / "patch10_output" / "pdf"))
        self.redact_terms_var = tk.StringVar(value="SECRET")
        self.pdf_open_password_var = tk.StringVar()
        self.pdf_user_password_var = tk.StringVar()
        self.pdf_owner_password_var = tk.StringVar()
        self.pdf_compression_var = tk.StringVar(value="balanced")

        self.mail_sender_var = tk.StringVar()
        self.mail_to_var = tk.StringVar()
        self.mail_cc_var = tk.StringVar()
        self.mail_subject_var = tk.StringVar(value="Patch 10 Draft")
        self.mail_body_var = tk.StringVar(value="Hello from Patch 10")
        self.mail_eml_output_var = tk.StringVar(value=str(Path.cwd() / "patch10_output" / "mail" / "draft_message.eml"))
        self.mail_attachments_var = tk.StringVar()
        self.smtp_host_var = tk.StringVar()
        self.smtp_port_var = tk.StringVar(value="587")
        self.smtp_user_var = tk.StringVar()
        self.smtp_password_var = tk.StringVar()
        self.smtp_use_tls_var = tk.BooleanVar(value=True)
        self.smtp_use_ssl_var = tk.BooleanVar(value=False)

        self.status_var = tk.StringVar(value="Ready")

        self._build_ui()
        self.after(100, self._poll_queue)

    def _build_ui(self) -> None:
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        header = ttk.Frame(self, padding=(18, 16, 18, 10))
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        ttk.Label(header, text=APP_TITLE, font=("Segoe UI", 18, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Standalone Patch 10 feature pack: OCR, advanced PDF operations, and mail workflows for integration into your main branch.",
            wraplength=980,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        notebook = ttk.Notebook(self)
        notebook.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 12))

        ocr_tab = ttk.Frame(notebook, padding=14)
        pdf_tab = ttk.Frame(notebook, padding=14)
        mail_tab = ttk.Frame(notebook, padding=14)
        notebook.add(ocr_tab, text="OCR")
        notebook.add(pdf_tab, text="Advanced PDF")
        notebook.add(mail_tab, text="Mail")

        self._build_ocr_tab(ocr_tab)
        self._build_pdf_tab(pdf_tab)
        self._build_mail_tab(mail_tab)

        footer = ttk.Frame(self, padding=(18, 0, 18, 18))
        footer.grid(row=2, column=0, sticky="ew")
        footer.grid_columnconfigure(0, weight=1)
        ttk.Label(footer, textvariable=self.status_var).grid(row=0, column=0, sticky="w")

    def _build_ocr_tab(self, parent: ttk.Frame) -> None:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(2, weight=1)

        form = ttk.LabelFrame(parent, text="OCR inputs")
        form.grid(row=0, column=0, sticky="ew")
        form.grid_columnconfigure(1, weight=1)

        ttk.Label(form, text="Input image or PDF:").grid(row=0, column=0, sticky="w", padx=10, pady=(12, 6))
        ttk.Entry(form, textvariable=self.ocr_input_var).grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=(12, 6))
        ttk.Button(form, text="Browse", command=self._browse_ocr_input).grid(row=0, column=2, sticky="e", padx=(0, 10), pady=(12, 6))

        ttk.Label(form, text="Output folder:").grid(row=1, column=0, sticky="w", padx=10, pady=6)
        ttk.Entry(form, textvariable=self.ocr_output_var).grid(row=1, column=1, sticky="ew", padx=(0, 8), pady=6)
        ttk.Button(form, text="Browse", command=self._browse_ocr_output).grid(row=1, column=2, sticky="e", padx=(0, 10), pady=6)

        ttk.Label(form, text="Language:").grid(row=2, column=0, sticky="w", padx=10, pady=6)
        ttk.Entry(form, textvariable=self.ocr_lang_var, width=12).grid(row=2, column=1, sticky="w", pady=6)
        ttk.Label(form, text="DPI:").grid(row=2, column=2, sticky="e", pady=6)
        ttk.Entry(form, textvariable=self.ocr_dpi_var, width=8).grid(row=2, column=3, sticky="w", padx=(8, 10), pady=6)
        ttk.Label(form, text="PSM:").grid(row=2, column=4, sticky="e", pady=6)
        ttk.Entry(form, textvariable=self.ocr_psm_var, width=8).grid(row=2, column=5, sticky="w", padx=(8, 10), pady=6)

        actions = ttk.Frame(parent)
        actions.grid(row=1, column=0, sticky="ew", pady=(12, 10))
        ttk.Button(actions, text="Image -> Searchable PDF", command=self._run_image_to_searchable_pdf).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(actions, text="PDF -> Searchable PDF", command=self._run_pdf_to_searchable_pdf).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(actions, text="Extract OCR Text", command=self._run_extract_ocr_text).grid(row=0, column=2)

        self.ocr_log = ScrolledText(parent, wrap=tk.WORD, height=22)
        self.ocr_log.grid(row=2, column=0, sticky="nsew")
        self.ocr_log.configure(state="disabled")

    def _build_pdf_tab(self, parent: ttk.Frame) -> None:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(2, weight=1)

        form = ttk.LabelFrame(parent, text="PDF inputs")
        form.grid(row=0, column=0, sticky="ew")
        form.grid_columnconfigure(1, weight=1)

        ttk.Label(form, text="Input PDF:").grid(row=0, column=0, sticky="w", padx=10, pady=(12, 6))
        ttk.Entry(form, textvariable=self.pdf_input_var).grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=(12, 6))
        ttk.Button(form, text="Browse", command=self._browse_pdf_input).grid(row=0, column=2, sticky="e", padx=(0, 10), pady=(12, 6))

        ttk.Label(form, text="Output folder:").grid(row=1, column=0, sticky="w", padx=10, pady=6)
        ttk.Entry(form, textvariable=self.pdf_output_var).grid(row=1, column=1, sticky="ew", padx=(0, 8), pady=6)
        ttk.Button(form, text="Browse", command=self._browse_pdf_output).grid(row=1, column=2, sticky="e", padx=(0, 10), pady=6)

        ttk.Label(form, text="Redaction terms:").grid(row=2, column=0, sticky="w", padx=10, pady=6)
        ttk.Entry(form, textvariable=self.redact_terms_var).grid(row=2, column=1, sticky="ew", padx=(0, 8), pady=6)
        ttk.Label(form, text="Open/current password:").grid(row=3, column=0, sticky="w", padx=10, pady=6)
        ttk.Entry(form, textvariable=self.pdf_open_password_var, show="*").grid(row=3, column=1, sticky="ew", padx=(0, 8), pady=6)

        ttk.Label(form, text="User password:").grid(row=4, column=0, sticky="w", padx=10, pady=6)
        ttk.Entry(form, textvariable=self.pdf_user_password_var, show="*").grid(row=4, column=1, sticky="ew", padx=(0, 8), pady=6)
        ttk.Label(form, text="Owner password:").grid(row=5, column=0, sticky="w", padx=10, pady=6)
        ttk.Entry(form, textvariable=self.pdf_owner_password_var, show="*").grid(row=5, column=1, sticky="ew", padx=(0, 8), pady=6)

        ttk.Label(form, text="Compression profile:").grid(row=6, column=0, sticky="w", padx=10, pady=(6, 12))
        ttk.Combobox(form, textvariable=self.pdf_compression_var, values=["safe", "balanced", "strong"], state="readonly", width=14).grid(
            row=6, column=1, sticky="w", pady=(6, 12)
        )

        actions = ttk.Frame(parent)
        actions.grid(row=1, column=0, sticky="ew", pady=(12, 10))
        ttk.Button(actions, text="Redact Text", command=self._run_redact).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(actions, text="Lock PDF", command=self._run_lock_pdf).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(actions, text="Unlock PDF", command=self._run_unlock_pdf).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(actions, text="Compress PDF", command=self._run_compress_pdf).grid(row=0, column=3)

        self.pdf_log = ScrolledText(parent, wrap=tk.WORD, height=22)
        self.pdf_log.grid(row=2, column=0, sticky="nsew")
        self.pdf_log.configure(state="disabled")

    def _build_mail_tab(self, parent: ttk.Frame) -> None:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(2, weight=1)

        form = ttk.LabelFrame(parent, text="Mail setup")
        form.grid(row=0, column=0, sticky="ew")
        form.grid_columnconfigure(1, weight=1)

        fields = [
            ("Sender:", self.mail_sender_var),
            ("To (comma separated):", self.mail_to_var),
            ("Cc (comma separated):", self.mail_cc_var),
            ("Subject:", self.mail_subject_var),
            ("EML output path:", self.mail_eml_output_var),
            ("Attachments (comma separated):", self.mail_attachments_var),
            ("SMTP host:", self.smtp_host_var),
            ("SMTP port:", self.smtp_port_var),
            ("SMTP username:", self.smtp_user_var),
            ("SMTP password:", self.smtp_password_var),
        ]
        for row, (label, variable) in enumerate(fields):
            ttk.Label(form, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=(10 if row == 0 else 6, 6))
            show = "*" if label == "SMTP password:" else None
            entry = ttk.Entry(form, textvariable=variable, show=show or "")
            entry.grid(row=row, column=1, sticky="ew", padx=(0, 8), pady=(10 if row == 0 else 6, 6))
            if label == "EML output path:":
                ttk.Button(form, text="Browse", command=self._browse_eml_output).grid(row=row, column=2, sticky="e", padx=(0, 10), pady=6)
            elif label == "Attachments (comma separated):":
                ttk.Button(form, text="Browse", command=self._browse_attachments).grid(row=row, column=2, sticky="e", padx=(0, 10), pady=6)

        opts = ttk.Frame(form)
        opts.grid(row=len(fields), column=0, columnspan=3, sticky="w", padx=10, pady=(10, 12))
        ttk.Checkbutton(opts, text="Use STARTTLS", variable=self.smtp_use_tls_var).grid(row=0, column=0, padx=(0, 10))
        ttk.Checkbutton(opts, text="Use SSL", variable=self.smtp_use_ssl_var).grid(row=0, column=1)

        ttk.Label(parent, text="Body:").grid(row=1, column=0, sticky="w")
        self.mail_body = ScrolledText(parent, wrap=tk.WORD, height=10)
        self.mail_body.grid(row=2, column=0, sticky="nsew", pady=(6, 10))
        self.mail_body.insert("1.0", self.mail_body_var.get())

        actions = ttk.Frame(parent)
        actions.grid(row=3, column=0, sticky="ew")
        ttk.Button(actions, text="Create EML Draft", command=self._run_build_eml).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(actions, text="Open mailto Draft", command=self._run_mailto).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(actions, text="Send via SMTP", command=self._run_send_smtp).grid(row=0, column=2)

        self.mail_log = ScrolledText(parent, wrap=tk.WORD, height=10)
        self.mail_log.grid(row=4, column=0, sticky="nsew", pady=(10, 0))
        self.mail_log.configure(state="disabled")

    def _append_log(self, widget: ScrolledText, message: str) -> None:
        widget.configure(state="normal")
        widget.insert(tk.END, message.rstrip() + "\n")
        widget.see(tk.END)
        widget.configure(state="disabled")

    def _run_async(self, label: str, func, *, widget: ScrolledText) -> None:
        if self.running:
            return
        self.running = True
        self.status_var.set(f"Running: {label}")
        self._append_log(widget, f"=== {label} ===")

        def worker() -> None:
            try:
                result = func()
                self.worker_queue.put(("done", (widget, label, result)))
            except Exception as exc:
                self.worker_queue.put(("error", (widget, label, exc)))

        threading.Thread(target=worker, daemon=True).start()

    def _poll_queue(self) -> None:
        try:
            while True:
                event, payload = self.worker_queue.get_nowait()
                widget, label, data = payload
                if event == "done":
                    self._append_log(widget, f"Done: {label}")
                    if data is not None:
                        self._append_log(widget, str(data))
                    self.status_var.set(f"Completed: {label}")
                else:
                    self._append_log(widget, f"Error: {label}")
                    self._append_log(widget, str(data))
                    self.status_var.set(f"Failed: {label}")
                    messagebox.showerror(APP_TITLE, str(data))
                self.running = False
        except queue.Empty:
            pass
        finally:
            self.after(120, self._poll_queue)

    def _ocr_config(self) -> OcrConfig:
        return OcrConfig(
            language=self.ocr_lang_var.get().strip() or "eng",
            dpi=int(self.ocr_dpi_var.get().strip() or "220"),
            psm=int(self.ocr_psm_var.get().strip() or "6"),
        )

    def _ocr_output_path(self, suffix: str) -> Path:
        out_dir = Path(self.ocr_output_var.get().strip() or Path.cwd() / "patch10_output" / "ocr")
        out_dir.mkdir(parents=True, exist_ok=True)
        source = Path(self.ocr_input_var.get().strip())
        return out_dir / f"{source.stem}_{suffix}"

    def _pdf_output_path(self, suffix: str) -> Path:
        out_dir = Path(self.pdf_output_var.get().strip() or Path.cwd() / "patch10_output" / "pdf")
        out_dir.mkdir(parents=True, exist_ok=True)
        source = Path(self.pdf_input_var.get().strip())
        return out_dir / f"{source.stem}_{suffix}"

    def _browse_ocr_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Select image or PDF",
            filetypes=[("Image or PDF", "*.png *.jpg *.jpeg *.tif *.tiff *.webp *.bmp *.pdf"), ("All files", "*.*")],
        )
        if path:
            self.ocr_input_var.set(path)

    def _browse_ocr_output(self) -> None:
        path = filedialog.askdirectory(title="Select OCR output folder")
        if path:
            self.ocr_output_var.set(path)

    def _browse_pdf_input(self) -> None:
        path = filedialog.askopenfilename(title="Select PDF", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if path:
            self.pdf_input_var.set(path)

    def _browse_pdf_output(self) -> None:
        path = filedialog.askdirectory(title="Select PDF output folder")
        if path:
            self.pdf_output_var.set(path)

    def _browse_eml_output(self) -> None:
        path = filedialog.asksaveasfilename(title="Save EML draft", defaultextension=".eml", filetypes=[("EML message", "*.eml"), ("All files", "*.*")])
        if path:
            self.mail_eml_output_var.set(path)

    def _browse_attachments(self) -> None:
        paths = filedialog.askopenfilenames(title="Select attachments")
        if paths:
            self.mail_attachments_var.set(", ".join(paths))

    def _collect_attachments(self) -> list[Path]:
        value = self.mail_attachments_var.get().strip()
        if not value:
            return []
        return [Path(item.strip()) for item in value.split(",") if item.strip()]

    def _run_image_to_searchable_pdf(self) -> None:
        def job() -> Path:
            source = Path(self.ocr_input_var.get().strip())
            return image_to_searchable_pdf(source, self._ocr_output_path("searchable.pdf"), config=self._ocr_config())
        self._run_async("Image -> Searchable PDF", job, widget=self.ocr_log)

    def _run_pdf_to_searchable_pdf(self) -> None:
        def job() -> Path:
            source = Path(self.ocr_input_var.get().strip())
            return pdf_to_searchable_pdf(source, self._ocr_output_path("ocr.pdf"), config=self._ocr_config())
        self._run_async("PDF -> Searchable PDF", job, widget=self.ocr_log)

    def _run_extract_ocr_text(self) -> None:
        def job() -> Path:
            source = Path(self.ocr_input_var.get().strip())
            return extract_text_with_ocr(source, self._ocr_output_path("ocr.txt"), config=self._ocr_config())
        self._run_async("Extract OCR Text", job, widget=self.ocr_log)

    def _run_redact(self) -> None:
        def job() -> Path:
            source = Path(self.pdf_input_var.get().strip())
            phrases = [part.strip() for part in self.redact_terms_var.get().split(";") if part.strip()]
            return redact_text(source, self._pdf_output_path("redacted.pdf"), phrases)
        self._run_async("Redact Text", job, widget=self.pdf_log)

    def _run_lock_pdf(self) -> None:
        def job() -> Path:
            source = Path(self.pdf_input_var.get().strip())
            return password_protect_pdf(
                source,
                self._pdf_output_path("locked.pdf"),
                user_password=self.pdf_user_password_var.get().strip(),
                owner_password=self.pdf_owner_password_var.get().strip() or None,
            )
        self._run_async("Lock PDF", job, widget=self.pdf_log)

    def _run_unlock_pdf(self) -> None:
        def job() -> Path:
            source = Path(self.pdf_input_var.get().strip())
            return remove_pdf_password(
                source,
                self._pdf_output_path("unlocked.pdf"),
                password=self.pdf_open_password_var.get().strip(),
            )
        self._run_async("Unlock PDF", job, widget=self.pdf_log)

    def _run_compress_pdf(self) -> None:
        def job() -> Path:
            source = Path(self.pdf_input_var.get().strip())
            return compress_pdf(
                source,
                self._pdf_output_path("compressed.pdf"),
                profile=self.pdf_compression_var.get().strip(),
                password=self.pdf_open_password_var.get().strip(),
            )
        self._run_async("Compress PDF", job, widget=self.pdf_log)

    def _build_smtp_config(self) -> SmtpConfig:
        return SmtpConfig(
            host=self.smtp_host_var.get().strip(),
            port=int(self.smtp_port_var.get().strip() or "587"),
            username=self.smtp_user_var.get().strip(),
            password=self.smtp_password_var.get().strip(),
            sender=self.mail_sender_var.get().strip(),
            use_tls=bool(self.smtp_use_tls_var.get()),
            use_ssl=bool(self.smtp_use_ssl_var.get()),
        )

    def _mail_recipients(self) -> tuple[list[str], list[str]]:
        to = [item.strip() for item in self.mail_to_var.get().split(",") if item.strip()]
        cc = [item.strip() for item in self.mail_cc_var.get().split(",") if item.strip()]
        return to, cc

    def _run_build_eml(self) -> None:
        def job() -> Path:
            to, cc = self._mail_recipients()
            return build_eml_draft(
                self.mail_eml_output_var.get().strip(),
                sender=self.mail_sender_var.get().strip(),
                to=to,
                cc=cc,
                subject=self.mail_subject_var.get().strip(),
                body=self.mail_body.get("1.0", tk.END).strip(),
                attachments=self._collect_attachments(),
            )
        self._run_async("Create EML Draft", job, widget=self.mail_log)

    def _run_mailto(self) -> None:
        def job() -> str:
            to, cc = self._mail_recipients()
            open_mailto_draft(
                to=to,
                cc=cc,
                subject=self.mail_subject_var.get().strip(),
                body=self.mail_body.get("1.0", tk.END).strip(),
            )
            return "Opened the default mail client using a mailto draft."
        self._run_async("Open mailto Draft", job, widget=self.mail_log)

    def _run_send_smtp(self) -> None:
        def job() -> str:
            to, cc = self._mail_recipients()
            send_email_smtp(
                self._build_smtp_config(),
                to=to,
                cc=cc,
                subject=self.mail_subject_var.get().strip(),
                body=self.mail_body.get("1.0", tk.END).strip(),
                attachments=self._collect_attachments(),
            )
            return "SMTP send completed successfully."
        self._run_async("Send via SMTP", job, widget=self.mail_log)


def main() -> None:
    parser = argparse.ArgumentParser(description=APP_TITLE)
    parser.add_argument("--smoke-test-ui", action="store_true", help="Open the GUI briefly and exit.")
    args = parser.parse_args()

    app = Patch10DemoApp()
    if args.smoke_test_ui:
        app.after(1000, app.destroy)
    app.mainloop()


if __name__ == "__main__":
    main()
