import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import mammoth
from markdownify import markdownify as md


class ConverterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("DOCX 转换工具")
        self.root.geometry("760x470")

        self.docx_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.strict_mode = tk.BooleanVar(value=False)

        self._build_ui()

    def _build_ui(self) -> None:
        padding = {"padx": 10, "pady": 6}

        main = ttk.Frame(self.root)
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main, text="DOCX 文件:").grid(row=0, column=0, sticky=tk.W, **padding)
        ttk.Entry(main, textvariable=self.docx_path, width=70).grid(
            row=0, column=1, sticky=tk.EW, **padding
        )
        ttk.Button(main, text="选择...", command=self.pick_docx).grid(
            row=0, column=2, sticky=tk.E, **padding
        )

        ttk.Label(main, text="输出目录:").grid(row=1, column=0, sticky=tk.W, **padding)
        ttk.Entry(main, textvariable=self.output_dir, width=70).grid(
            row=1, column=1, sticky=tk.EW, **padding
        )
        ttk.Button(main, text="选择...", command=self.pick_output_dir).grid(
            row=1, column=2, sticky=tk.E, **padding
        )

        button_bar = ttk.Frame(main)
        button_bar.grid(row=2, column=0, columnspan=3, sticky=tk.W, padx=10, pady=8)

        self.btn_md = ttk.Button(button_bar, text="转 Markdown", command=self.convert_md)
        self.btn_md.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_pdf = ttk.Button(button_bar, text="转 PDF", command=self.convert_pdf)
        self.btn_pdf.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_both = ttk.Button(button_bar, text="全部转换", command=self.convert_both)
        self.btn_both.pack(side=tk.LEFT)

        ttk.Checkbutton(
            main,
            text="严格模式（检测到内容风险时高亮提醒）",
            variable=self.strict_mode,
        ).grid(row=2, column=1, columnspan=2, sticky=tk.E, padx=10, pady=8)

        ttk.Label(main, text="日志:").grid(row=3, column=0, sticky=tk.NW, **padding)

        self.log_box = tk.Text(main, height=20, wrap=tk.WORD)
        self.log_box.grid(row=3, column=1, columnspan=2, sticky=tk.NSEW, padx=10, pady=6)

        scrollbar = ttk.Scrollbar(main, orient=tk.VERTICAL, command=self.log_box.yview)
        scrollbar.grid(row=3, column=3, sticky=tk.NS, pady=6)
        self.log_box.configure(yscrollcommand=scrollbar.set)

        main.columnconfigure(1, weight=1)
        main.rowconfigure(3, weight=1)

    def pick_docx(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 DOCX 文件", filetypes=[("Word 文件", "*.docx")]
        )
        if path:
            self.docx_path.set(path)
            if not self.output_dir.get():
                self.output_dir.set(str(Path(path).parent))

    def pick_output_dir(self) -> None:
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir.set(path)

    def append_log(self, text: str) -> None:
        self.log_box.insert(tk.END, text + "\n")
        self.log_box.see(tk.END)

    def set_buttons(self, enabled: bool) -> None:
        state = tk.NORMAL if enabled else tk.DISABLED
        self.btn_md.config(state=state)
        self.btn_pdf.config(state=state)
        self.btn_both.config(state=state)

    def validate_inputs(self) -> tuple[Path, Path] | tuple[None, None]:
        docx = self.docx_path.get().strip()
        out_dir = self.output_dir.get().strip()

        if not docx:
            messagebox.showerror("错误", "请先选择 DOCX 文件")
            return None, None
        if not out_dir:
            messagebox.showerror("错误", "请先选择输出目录")
            return None, None

        docx_path = Path(docx)
        out_path = Path(out_dir)

        if not docx_path.exists() or docx_path.suffix.lower() != ".docx":
            messagebox.showerror("错误", "DOCX 文件不存在或格式不正确")
            return None, None

        out_path.mkdir(parents=True, exist_ok=True)
        return docx_path, out_path

    def run_task(self, task) -> None:
        paths = self.validate_inputs()
        if not paths or paths[0] is None:
            return

        docx_path, out_path = paths
        self.set_buttons(False)

        def worker() -> None:
            try:
                task(docx_path, out_path)
            except Exception as exc:
                self.root.after(0, lambda: self.append_log(f"[失败] {exc}"))
                self.root.after(0, lambda: messagebox.showerror("转换失败", str(exc)))
            finally:
                self.root.after(0, lambda: self.set_buttons(True))

        threading.Thread(target=worker, daemon=True).start()

    def convert_md(self) -> None:
        self.run_task(self._convert_docx_to_md)

    def convert_pdf(self) -> None:
        self.run_task(self._convert_docx_to_pdf)

    def convert_both(self) -> None:
        def task(docx_path: Path, out_path: Path) -> None:
            self._convert_docx_to_md(docx_path, out_path)
            self._convert_docx_to_pdf(docx_path, out_path)

        self.run_task(task)

    def _convert_docx_to_md(self, docx_path: Path, out_path: Path) -> None:
        self.root.after(0, lambda: self.append_log(f"开始转 Markdown: {docx_path.name}"))

        stem = docx_path.stem
        image_dir = out_path / f"{stem}_images"
        image_dir.mkdir(parents=True, exist_ok=True)

        image_index = {"count": 0}

        def convert_image(image):
            image_index["count"] += 1
            ext = image.content_type.split("/")[-1].lower()
            if ext == "jpeg":
                ext = "jpg"
            filename = f"img_{image_index['count']:03d}.{ext}"
            target = image_dir / filename

            with image.open() as image_bytes:
                target.write_bytes(image_bytes.read())

            rel_path = f"{stem}_images/{filename}"
            return {"src": rel_path}

        with open(docx_path, "rb") as f:
            result = mammoth.convert_to_html(
                f, convert_image=mammoth.images.img_element(convert_image)
            )

        html = result.value
        md_text = md(html, heading_style="ATX")

        md_file = out_path / f"{stem}.md"
        md_file.write_text(md_text, encoding="utf-8")

        self._report_mammoth_messages(result.messages)
        self.root.after(0, lambda: self.append_log(f"[完成] Markdown: {md_file}"))

    def _report_mammoth_messages(self, messages) -> None:
        if not messages:
            return

        raw_messages = [m.message for m in messages]
        style_hints: list[str] = []
        content_risks: list[str] = []
        other_warnings: list[str] = []

        style_keywords = [
            "unrecognised paragraph style",
            "style id",
            "tblprex",
            "table style",
        ]
        content_keywords = [
            "image",
            "could not",
            "not supported",
            "unsupported",
            "failed",
            "error",
            "text box",
            "footnote",
            "endnote",
            "comment",
        ]

        for msg in raw_messages:
            lowered = msg.lower()
            if any(k in lowered for k in style_keywords):
                style_hints.append(msg)
            elif any(k in lowered for k in content_keywords):
                content_risks.append(msg)
            else:
                other_warnings.append(msg)

        if style_hints:
            self.root.after(
                0,
                lambda: self.append_log(
                    "[样式提示] 以下告警通常不影响正文内容：" + "; ".join(style_hints)
                ),
            )

        if content_risks:
            self.root.after(
                0, lambda: self.append_log("[内容风险] " + "; ".join(content_risks))
            )
            if self.strict_mode.get():
                self.root.after(
                    0,
                    lambda: messagebox.showwarning(
                        "严格模式提醒",
                        "检测到可能影响内容完整性的告警，请重点核对转换结果。\n\n"
                        + "\n".join(content_risks[:5]),
                    ),
                )

        if other_warnings:
            self.root.after(
                0,
                lambda: self.append_log("[一般警告] " + "; ".join(other_warnings)),
            )
            if self.strict_mode.get():
                self.root.after(
                    0,
                    lambda: messagebox.showwarning(
                        "严格模式提醒",
                        "检测到未分类告警，请人工复核。\n\n"
                        + "\n".join(other_warnings[:5]),
                    ),
                )

    def _convert_docx_to_pdf(self, docx_path: Path, out_path: Path) -> None:
        self.root.after(0, lambda: self.append_log(f"开始转 PDF: {docx_path.name}"))

        try:
            from docx2pdf import convert
        except ImportError as exc:
            raise RuntimeError("缺少依赖 docx2pdf，请先安装 requirements.txt") from exc

        pdf_file = out_path / f"{docx_path.stem}.pdf"

        try:
            convert(str(docx_path), str(pdf_file))
        except Exception as exc:
            raise RuntimeError(
                "DOCX 转 PDF 失败。Windows 下通常需要已安装 Microsoft Word。"
            ) from exc

        self.root.after(0, lambda: self.append_log(f"[完成] PDF: {pdf_file}"))


def main() -> None:
    root = tk.Tk()
    app = ConverterApp(root)
    app.append_log("准备就绪：请选择 DOCX 文件和输出目录。")
    root.mainloop()


if __name__ == "__main__":
    main()
