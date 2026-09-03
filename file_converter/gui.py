import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from . import registry
from .converters.base import ConvertError
from .converters.pdf import merge_pdfs, rotate_pdf, split_pdf

ENCODINGS = ("utf-8", "gbk", "big5", "shift_jis")


class ConverterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("万能文件转换工具")
        root.minsize(760, 520)
        self.files: list[Path] = []
        self.queue: queue.Queue = queue.Queue()

        top = ttk.Frame(root, padding=8)
        top.pack(fill=tk.BOTH, expand=True)

        bar = ttk.Frame(top)
        bar.pack(fill=tk.X)
        ttk.Button(bar, text="添加文件", command=self.add_files).pack(side=tk.LEFT)
        ttk.Button(bar, text="添加文件夹", command=self.add_folder).pack(side=tk.LEFT, padx=4)
        ttk.Button(bar, text="移除选中", command=self.remove_selected).pack(side=tk.LEFT)
        ttk.Button(bar, text="清空", command=self.clear_files).pack(side=tk.LEFT, padx=4)
        ttk.Separator(bar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8)
        ttk.Button(bar, text="合并PDF", command=self.merge_pdf).pack(side=tk.LEFT)
        ttk.Button(bar, text="拆分PDF", command=self.split_pdf).pack(side=tk.LEFT, padx=4)
        ttk.Button(bar, text="旋转PDF", command=self.rotate_pdf).pack(side=tk.LEFT)

        self.listbox = tk.Listbox(top, selectmode=tk.EXTENDED, activestyle="none")
        scroll = ttk.Scrollbar(top, orient=tk.VERTICAL, command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scroll.set)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=6)
        scroll.pack(side=tk.RIGHT, fill=tk.Y, pady=6)
        self.listbox.bind("<<ListboxSelect>>", lambda _e: self.refresh_targets())

        opts = ttk.Frame(root, padding=8)
        opts.pack(fill=tk.X)
        ttk.Label(opts, text="目标格式:").pack(side=tk.LEFT)
        self.target = ttk.Combobox(opts, width=10, state="readonly")
        self.target.pack(side=tk.LEFT, padx=4)
        ttk.Label(opts, text="输出目录:").pack(side=tk.LEFT, padx=(16, 0))
        self.outdir = ttk.Entry(opts, width=36)
        self.outdir.pack(side=tk.LEFT, padx=4)
        self.outdir.insert(0, "")
        ttk.Button(opts, text="浏览…", command=self.pick_outdir).pack(side=tk.LEFT)
        ttk.Label(opts, text="DPI:").pack(side=tk.LEFT, padx=(16, 0))
        self.dpi = ttk.Spinbox(opts, from_=72, to=600, width=6)
        self.dpi.set(150)
        self.dpi.pack(side=tk.LEFT, padx=4)
        ttk.Label(opts, text="文本编码:").pack(side=tk.LEFT, padx=(16, 0))
        self.encoding = ttk.Combobox(opts, width=10, state="readonly", values=ENCODINGS)
        self.encoding.set("utf-8")
        self.encoding.pack(side=tk.LEFT, padx=4)

        self.progress = ttk.Progressbar(root, mode="determinate")
        self.progress.pack(fill=tk.X, padx=8)
        self.log = tk.Text(root, height=8, state=tk.DISABLED)
        self.log.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)
        self.convert_btn = ttk.Button(root, text="开始转换", command=self.start_convert)
        self.convert_btn.pack(pady=(0, 8))

    # ---- 文件列表 ----
    def add_files(self):
        for f in filedialog.askopenfilenames(title="选择要转换的文件"):
            self._add(Path(f))

    def add_folder(self):
        folder = filedialog.askdirectory(title="选择文件夹")
        if folder:
            for p in sorted(Path(folder).rglob("*")):
                if p.is_file():
                    self._add(p)

    def _add(self, p: Path):
        if p not in self.files:
            self.files.append(p)
            self.listbox.insert(tk.END, str(p))
        self.refresh_targets()

    def remove_selected(self):
        for i in reversed(self.listbox.curselection()):
            self.listbox.delete(i)
            del self.files[i]
        self.refresh_targets()

    def clear_files(self):
        self.files.clear()
        self.listbox.delete(0, tk.END)
        self.refresh_targets()

    def refresh_targets(self):
        indices = self.listbox.curselection() or range(len(self.files))
        exts = {self.files[i].suffix.lower().lstrip(".") for i in indices if i < len(self.files)}
        common: set[str] | None = None
        for ext in exts:
            targets = set(registry.available_targets(ext))
            common = targets if common is None else common & targets
        values = sorted(common or [])
        self.target.configure(values=values)
        if values and self.target.get() not in values:
            self.target.set(values[0])

    def pick_outdir(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.outdir.delete(0, tk.END)
            self.outdir.insert(0, d)

    def log_line(self, text: str):
        self.log.configure(state=tk.NORMAL)
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)
        self.log.configure(state=tk.DISABLED)

    # ---- 转换 ----
    def start_convert(self):
        if not self.files:
            messagebox.showwarning("提示", "请先添加要转换的文件")
            return
        target = self.target.get()
        if not target:
            messagebox.showwarning("提示", "请选择目标格式")
            return
        outdir = self.outdir.get().strip() or None
        dpi = int(self.dpi.get() or 150)
        dst_encoding = self.encoding.get() or "utf-8"
        files = list(self.files)
        self.convert_btn.configure(state=tk.DISABLED)
        self.progress.configure(maximum=len(files), value=0)
        threading.Thread(
            target=self._worker,
            args=(files, target, outdir, dpi, dst_encoding),
            daemon=True,
        ).start()
        self.root.after(100, self._poll)

    def _worker(self, files, target, outdir, dpi, dst_encoding):
        ok, fail = 0, 0
        for src in files:
            try:
                dst = None
                if outdir:
                    dst = registry.unique_path(Path(outdir) / (src.stem + "." + target))
                out = registry.convert(src, target, dst, dpi=dpi, dst_encoding=dst_encoding)
                self.queue.put(("log", f"[成功] {src.name} → {out}"))
                ok += 1
            except ConvertError as e:
                self.queue.put(("log", f"[失败] {src.name}: {e}"))
                fail += 1
            self.queue.put(("step", None))
        self.queue.put(("done", (ok, fail)))

    def _poll(self):
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "log":
                    self.log_line(payload)
                elif kind == "step":
                    self.progress.step(1)
                elif kind == "done":
                    ok, fail = payload
                    self.convert_btn.configure(state=tk.NORMAL)
                    messagebox.showinfo("转换完成", f"成功 {ok} 个，失败 {fail} 个")
                    return
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    # ---- PDF 工具 ----
    def _selected_pdfs(self) -> list[Path]:
        idx = self.listbox.curselection()
        chosen = [self.files[i] for i in idx] if idx else list(self.files)
        return [p for p in chosen if p.suffix.lower() == ".pdf"]

    def merge_pdf(self):
        pdfs = self._selected_pdfs()
        if len(pdfs) < 2:
            messagebox.showwarning("提示", "请在列表中加入至少 2 个 PDF")
            return
        out = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not out:
            return
        try:
            merge_pdfs(pdfs, Path(out))
            self.log_line(f"[成功] 合并 → {out}")
        except ConvertError as e:
            messagebox.showerror("失败", str(e))

    def split_pdf(self):
        pdfs = self._selected_pdfs()
        if len(pdfs) != 1:
            messagebox.showwarning("提示", "拆分需要且只需要 1 个 PDF")
            return
        pages = simpledialog.askstring("拆分 PDF", "页码范围（如 1-3,5）：")
        if not pages:
            return
        outdir = filedialog.askdirectory(title="选择输出目录")
        if not outdir:
            return
        try:
            outs = split_pdf(pdfs[0], Path(outdir), pages)
            self.log_line("[成功] 拆分出 " + ", ".join(o.name for o in outs))
        except ConvertError as e:
            messagebox.showerror("失败", str(e))

    def rotate_pdf(self):
        pdfs = self._selected_pdfs()
        if len(pdfs) != 1:
            messagebox.showwarning("提示", "旋转需要且只需要 1 个 PDF")
            return
        angle = simpledialog.askinteger("旋转 PDF", "角度（90/180/270）：", initialvalue=90)
        if angle not in (90, 180, 270):
            return
        out = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not out:
            return
        try:
            rotate_pdf(pdfs[0], Path(out), angle)
            self.log_line(f"[成功] 旋转 {angle}° → {out}")
        except ConvertError as e:
            messagebox.showerror("失败", str(e))


def run():
    root = tk.Tk()
    ConverterApp(root)
    root.mainloop()
