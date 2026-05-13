from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


DEFAULT_DPI = 200
LANGUAGE_LABELS = {"zh": "中文", "en": "English"}
LABEL_TO_LANGUAGE = {label: code for code, label in LANGUAGE_LABELS.items()}

TEXT = {
    "zh": {
        "app_title": "PDF 转图片",
        "window_title": "PDF 页面导出为图片",
        "language": "语言",
        "pdf_file": "PDF 文件",
        "output_folder": "输出文件夹",
        "dpi": "清晰度 DPI",
        "pages": "导出页码",
        "pages_hint": "留空 = 全部，例如 1,3,5-8",
        "browse": "选择",
        "export": "开始导出",
        "initial_status": "请选择 PDF 文件和输出文件夹。",
        "opening_pdf": "正在打开 PDF...",
        "export_failed": "导出失败。",
        "valid_pdf": "请选择有效的 PDF 文件。",
        "create_output_failed": "无法创建输出文件夹：\n{error}",
        "valid_output": "请选择有效的输出文件夹。",
        "dpi_number": "DPI 必须是数字。",
        "dpi_range": "DPI 必须在 72 到 600 之间。",
        "missing_pymupdf": "缺少依赖：PyMuPDF。\n请先运行：pip install pymupdf",
        "progress": "已导出 PDF 第 {page} 页（{current}/{total}）。",
        "done_status": "完成。已导出 {count} 页到：{output}",
        "done_dialog": "已导出 {count} 页。",
        "select_pdf_title": "选择 PDF",
        "select_output_title": "选择输出文件夹",
        "pdf_files": "PDF 文件",
        "all_files": "所有文件",
        "empty_page_item": "页码选择中包含空项。",
        "invalid_page_range": "页码范围格式错误：{value}",
        "range_low_to_high": "页码范围必须从小到大：{value}",
        "invalid_page_number": "页码格式错误：{value}",
        "page_outside_range": "页码 {page} 超出范围 1-{page_count}。",
        "empty_pdf": "这个 PDF 没有页面。",
    },
    "en": {
        "app_title": "PDF to Images",
        "window_title": "Export PDF Pages as Images",
        "language": "Language",
        "pdf_file": "PDF file",
        "output_folder": "Output folder",
        "dpi": "DPI",
        "pages": "Pages",
        "pages_hint": "Blank = all, e.g. 1,3,5-8",
        "browse": "Browse",
        "export": "Export Pages",
        "initial_status": "Select a PDF file and an output folder.",
        "opening_pdf": "Opening PDF...",
        "export_failed": "Export failed.",
        "valid_pdf": "Please select a valid PDF file.",
        "create_output_failed": "Could not create output folder:\n{error}",
        "valid_output": "Please select a valid output folder.",
        "dpi_number": "DPI must be a number.",
        "dpi_range": "DPI must be between 72 and 600.",
        "missing_pymupdf": "Missing dependency: PyMuPDF.\nInstall it with: pip install pymupdf",
        "progress": "Exported PDF page {page} ({current} of {total}).",
        "done_status": "Done. Exported {count} pages to: {output}",
        "done_dialog": "Exported {count} pages.",
        "select_pdf_title": "Select PDF",
        "select_output_title": "Select output folder",
        "pdf_files": "PDF files",
        "all_files": "All files",
        "empty_page_item": "Page selection contains an empty item.",
        "invalid_page_range": "Invalid page range: {value}",
        "range_low_to_high": "Page range must go from low to high: {value}",
        "invalid_page_number": "Invalid page number: {value}",
        "page_outside_range": "Page {page} is outside 1-{page_count}.",
        "empty_pdf": "This PDF has no pages.",
    },
}


class AppError(ValueError):
    def __init__(self, message_key: str, **params: object) -> None:
        super().__init__(message_key)
        self.message_key = message_key
        self.params = params


def translate(language: str, key: str, **params: object) -> str:
    return TEXT[language][key].format(**params)


def parse_page_selection(page_spec: str, page_count: int) -> list[int]:
    spec = page_spec.strip()
    if not spec:
        return list(range(page_count))

    selected_pages: list[int] = []
    seen: set[int] = set()

    for raw_part in spec.split(","):
        part = raw_part.strip()
        if not part:
            raise AppError("empty_page_item")

        if "-" in part:
            range_parts = [item.strip() for item in part.split("-", 1)]
            if len(range_parts) != 2 or not range_parts[0] or not range_parts[1]:
                raise AppError("invalid_page_range", value=part)

            try:
                start = int(range_parts[0])
                end = int(range_parts[1])
            except ValueError as exc:
                raise AppError("invalid_page_range", value=part) from exc

            if start > end:
                raise AppError("range_low_to_high", value=part)

            page_numbers = range(start, end + 1)
        else:
            try:
                page = int(part)
            except ValueError as exc:
                raise AppError("invalid_page_number", value=part) from exc

            page_numbers = range(page, page + 1)

        for page_number in page_numbers:
            if page_number < 1 or page_number > page_count:
                raise AppError(
                    "page_outside_range",
                    page=page_number,
                    page_count=page_count,
                )

            page_index = page_number - 1
            if page_index not in seen:
                seen.add(page_index)
                selected_pages.append(page_index)

    return selected_pages


def export_pdf_pages_to_images(
    pdf: Path,
    output: Path,
    dpi: int,
    page_spec: str,
    progress_callback=None,
) -> int:
    import fitz  # PyMuPDF

    document = fitz.open(pdf)
    try:
        page_count = document.page_count

        if page_count == 0:
            raise AppError("empty_pdf")

        selected_pages = parse_page_selection(page_spec, page_count)
        scale = dpi / 72
        matrix = fitz.Matrix(scale, scale)
        total_selected = len(selected_pages)

        for export_index, page_index in enumerate(selected_pages, start=1):
            page = document.load_page(page_index)
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)
            image_path = output / f"{page_index + 1}.png"
            pixmap.save(image_path)

            if progress_callback:
                percent = int((export_index / total_selected) * 100)
                progress_callback(percent, page_index + 1, export_index, total_selected)

        return total_selected
    finally:
        document.close()


class PdfToImagesApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.language = "zh"
        self.title(translate(self.language, "app_title"))
        self.geometry("720x430")
        self.minsize(660, 390)

        self.pdf_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.dpi = tk.IntVar(value=DEFAULT_DPI)
        self.pages = tk.StringVar()
        self.language_label = tk.StringVar(value=LANGUAGE_LABELS[self.language])
        self.status_key = "initial_status"
        self.status_params: dict[str, object] = {}
        self.status = tk.StringVar(value=self.tr(self.status_key))
        self.progress = tk.IntVar(value=0)
        self.text_widgets: dict[str, list[tk.Widget]] = {}

        self._build_ui()
        self.language_label.trace_add("write", self.on_language_change)

    def tr(self, key: str, **params: object) -> str:
        return translate(self.language, key, **params)

    def set_status(self, key: str, **params: object) -> None:
        self.status_key = key
        self.status_params = params
        self.status.set(self.tr(key, **params))

    def register_text(self, key: str, widget: tk.Widget) -> None:
        self.text_widgets.setdefault(key, []).append(widget)
        widget.configure(text=self.tr(key))

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=18)
        outer.pack(fill=tk.BOTH, expand=True)
        outer.columnconfigure(1, weight=1)

        title = ttk.Label(outer, font=("Segoe UI", 16, "bold"))
        title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 18))
        self.register_text("window_title", title)

        language_frame = ttk.Frame(outer)
        language_frame.grid(row=0, column=2, sticky="e", pady=(0, 18))

        language_label = ttk.Label(language_frame)
        language_label.pack(side=tk.LEFT, padx=(0, 8))
        self.register_text("language", language_label)

        self.language_box = ttk.Combobox(
            language_frame,
            textvariable=self.language_label,
            values=list(LANGUAGE_LABELS.values()),
            state="readonly",
            width=10,
        )
        self.language_box.pack(side=tk.LEFT)

        pdf_label = ttk.Label(outer)
        pdf_label.grid(row=1, column=0, sticky="w", pady=6)
        self.register_text("pdf_file", pdf_label)

        pdf_entry = ttk.Entry(outer, textvariable=self.pdf_path)
        pdf_entry.grid(row=1, column=1, sticky="ew", padx=8, pady=6)

        pdf_button = ttk.Button(outer, command=self.pick_pdf)
        pdf_button.grid(row=1, column=2, pady=6)
        self.register_text("browse", pdf_button)

        output_label = ttk.Label(outer)
        output_label.grid(row=2, column=0, sticky="w", pady=6)
        self.register_text("output_folder", output_label)

        output_entry = ttk.Entry(outer, textvariable=self.output_dir)
        output_entry.grid(row=2, column=1, sticky="ew", padx=8, pady=6)

        output_button = ttk.Button(outer, command=self.pick_output_dir)
        output_button.grid(row=2, column=2, pady=6)
        self.register_text("browse", output_button)

        dpi_label = ttk.Label(outer)
        dpi_label.grid(row=3, column=0, sticky="w", pady=6)
        self.register_text("dpi", dpi_label)

        dpi_box = ttk.Spinbox(
            outer,
            from_=72,
            to=600,
            increment=25,
            textvariable=self.dpi,
            width=8,
        )
        dpi_box.grid(row=3, column=1, sticky="w", padx=8, pady=6)

        pages_label = ttk.Label(outer)
        pages_label.grid(row=4, column=0, sticky="w", pady=6)
        self.register_text("pages", pages_label)

        pages_entry = ttk.Entry(outer, textvariable=self.pages)
        pages_entry.grid(row=4, column=1, sticky="ew", padx=8, pady=6)

        pages_hint = ttk.Label(outer)
        pages_hint.grid(row=4, column=2, sticky="w", pady=6)
        self.register_text("pages_hint", pages_hint)

        self.progress_bar = ttk.Progressbar(
            outer,
            maximum=100,
            variable=self.progress,
            mode="determinate",
        )
        self.progress_bar.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(20, 8))

        status_label = ttk.Label(outer, textvariable=self.status, wraplength=660)
        status_label.grid(row=6, column=0, columnspan=3, sticky="w", pady=(0, 16))

        button_row = ttk.Frame(outer)
        button_row.grid(row=7, column=0, columnspan=3, sticky="e")

        self.convert_button = ttk.Button(button_row, command=self.start_export)
        self.convert_button.pack(side=tk.RIGHT)
        self.register_text("export", self.convert_button)

    def on_language_change(self, *_args: object) -> None:
        selected_language = LABEL_TO_LANGUAGE.get(self.language_label.get())
        if not selected_language or selected_language == self.language:
            return

        self.language = selected_language
        self.title(self.tr("app_title"))
        for key, widgets in self.text_widgets.items():
            for widget in widgets:
                widget.configure(text=self.tr(key))
        self.status.set(self.tr(self.status_key, **self.status_params))

    def show_error(self, message_key: str, **params: object) -> None:
        messagebox.showerror(self.tr("app_title"), self.tr(message_key, **params))

    def format_error(self, exc: Exception) -> str:
        if isinstance(exc, AppError):
            return self.tr(exc.message_key, **exc.params)
        return str(exc)

    def pick_pdf(self) -> None:
        filename = filedialog.askopenfilename(
            title=self.tr("select_pdf_title"),
            filetypes=[
                (self.tr("pdf_files"), "*.pdf"),
                (self.tr("all_files"), "*.*"),
            ],
        )
        if filename:
            self.pdf_path.set(filename)

    def pick_output_dir(self) -> None:
        directory = filedialog.askdirectory(title=self.tr("select_output_title"))
        if directory:
            self.output_dir.set(directory)

    def start_export(self) -> None:
        pdf = Path(self.pdf_path.get().strip())
        output = Path(self.output_dir.get().strip())

        if not pdf.is_file() or pdf.suffix.lower() != ".pdf":
            self.show_error("valid_pdf")
            return

        if not output.exists():
            try:
                output.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                self.show_error("create_output_failed", error=exc)
                return

        if not output.is_dir():
            self.show_error("valid_output")
            return

        try:
            dpi = int(self.dpi.get())
        except (tk.TclError, ValueError):
            self.show_error("dpi_number")
            return

        if dpi < 72 or dpi > 600:
            self.show_error("dpi_range")
            return

        self.progress.set(0)
        self.set_status("opening_pdf")
        self.convert_button.configure(state=tk.DISABLED)

        worker = threading.Thread(
            target=self.export_pdf_pages,
            args=(pdf, output, dpi, self.pages.get()),
            daemon=True,
        )
        worker.start()

    def export_pdf_pages(self, pdf: Path, output: Path, dpi: int, page_spec: str) -> None:
        try:
            import fitz  # noqa: F401
        except ImportError:
            self.after(0, self.fail_export, self.tr("missing_pymupdf"))
            return

        try:
            def on_progress(
                percent: int,
                pdf_page_number: int,
                export_index: int,
                total_selected: int,
            ) -> None:
                self.after(
                    0,
                    self.update_progress,
                    percent,
                    "progress",
                    {
                        "page": pdf_page_number,
                        "current": export_index,
                        "total": total_selected,
                    },
                )

            total_selected = export_pdf_pages_to_images(pdf, output, dpi, page_spec, on_progress)
            self.after(0, self.finish_export, total_selected, output)
        except Exception as exc:
            self.after(0, self.fail_export, self.format_error(exc))

    def update_progress(self, percent: int, status_key: str, params: dict[str, object]) -> None:
        self.progress.set(percent)
        self.set_status(status_key, **params)

    def finish_export(self, page_count: int, output: Path) -> None:
        self.convert_button.configure(state=tk.NORMAL)
        self.progress.set(100)
        self.set_status("done_status", count=page_count, output=output)
        messagebox.showinfo(self.tr("app_title"), self.tr("done_dialog", count=page_count))

    def fail_export(self, error: str) -> None:
        self.convert_button.configure(state=tk.NORMAL)
        self.set_status("export_failed")
        messagebox.showerror(self.tr("app_title"), error)


if __name__ == "__main__":
    app = PdfToImagesApp()
    app.mainloop()
