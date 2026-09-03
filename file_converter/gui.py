import queue
import threading
import tkinter as tk
import tkinter.font as tkfont
from pathlib import Path
from tkinter import filedialog, simpledialog, ttk

from . import registry
from .converters.base import ConvertError
from .converters.pdf import merge_pdfs, rotate_pdf, split_pdf

try:
    import tkinterdnd2

    HAS_DND = True
except ImportError:  # 拖放为可选增强，未安装时退化为点击添加
    tkinterdnd2 = None
    HAS_DND = False

ENCODINGS = ("utf-8", "gbk", "big5", "utf-16")

# 设计稿 index.html :root 调色板
BG = "#f5f5f5"
PANEL = "#ffffff"
TEXT = "#1a1a1a"
MUTED = "#5d5b5b"
FAINT = "#a0a0a0"
LINE = "#c4c4c4"
LINE_SOFT = "#e3e1dc"
ACCENT = "#778652"
ACCENT_DEEP = "#5c6a3f"
ROW_HOVER = "#fafaf7"
DRAG_BG = "#fbfbf9"
MENU_HOVER = "#f2f1ed"

FONT = "Microsoft YaHei"
MONO = "Consolas"

ICON_PATH = Path(__file__).with_name("assets") / "icon.png"

STATE_TEXT = {"idle": "等待", "doing": "转换中", "done": "完成", "fail": "失败"}


def _fmt_size(n: int) -> str:
    if n < 1024:
        return f"{n} B"
    if n < 1048576:
        return f"{n / 1024:.1f} KB"
    if n < 1073741824:
        return f"{n / 1048576:.1f} MB"
    return f"{n / 1073741824:.2f} GB"


def _round_rect(canvas, x1, y1, x2, y2, r, **kw):
    pts = [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]
    return canvas.create_polygon(pts, smooth=True, splinesteps=24, **kw)


class ConverterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("万能文件转换工具")
        root.geometry("480x620")
        root.minsize(440, 560)
        root.configure(bg=BG)
        try:
            if ICON_PATH.exists():
                self._icon = tk.PhotoImage(file=str(ICON_PATH))
                root.iconphoto(True, self._icon)
        except tk.TclError:
            pass

        self.files: list[dict] = []  # {path, state, err}
        self.queue: queue.Queue = queue.Queue()
        self.running = False
        self._drag_over = False
        self._adv_open = False
        self._links: list[tuple[tk.Label, tkfont.Font, object]] = []

        f = self._f = {
            "link": tkfont.Font(root=root, family=FONT, size=-12),
            "count": tkfont.Font(root=root, family=FONT, size=-11),
            "name": tkfont.Font(root=root, family=FONT, size=-12),
            "meta": tkfont.Font(root=root, family=FONT, size=-10),
            "mono": tkfont.Font(root=root, family=MONO, size=-10),
            "ext": tkfont.Font(root=root, family=MONO, size=-8),
            "label": tkfont.Font(root=root, family=FONT, size=-9),
            "empty": tkfont.Font(root=root, family=FONT, size=-12),
        }

        style = ttk.Style(root)
        style.configure(
            "Flat.TCombobox",
            fieldbackground=BG, background=BG, foreground=TEXT,
            borderwidth=0, relief="flat", padding=2, arrowsize=10,
        )
        style.map("Flat.TCombobox", fieldbackground=[("readonly", BG)],
                  foreground=[("readonly", TEXT)])
        try:
            style.layout("Flat.TCombobox", [
                ("Combobox.field", {"sticky": "nswe", "children": [
                    ("Combobox.downarrow", {"side": "right", "sticky": ""}),
                    ("Combobox.padding", {"expand": "1", "sticky": "nswe", "children": [
                        ("Combobox.focus", {"expand": "1", "sticky": "nswe", "children": [
                            ("Combobox.textarea", {"sticky": "nswe"})]})]})]})])
        except tk.TclError:
            pass
        style.configure(
            "Slim.Vertical.TScrollbar",
            background=LINE_SOFT, troughcolor=PANEL, bordercolor=PANEL,
            arrowcolor=PANEL, width=6, arrowsize=6,
        )

        self._build_toolbar()
        self._build_footer()  # 先 pack 底部，保证主区占剩余空间
        self._build_dropzone()
        self._render_rows()
        self._set_status("就绪")

    # ---- 界面骨架 ----
    def _make_link(self, parent, text, command, side=tk.LEFT, padx=(0, 16)):
        font = tkfont.Font(root=self.root, family=FONT, size=-12)
        lbl = tk.Label(parent, text=text, font=font, bg=BG, fg=TEXT,
                       cursor="hand2", padx=0, pady=0)
        lbl.pack(side=side, padx=padx)
        lbl.bind("<Enter>", lambda _e: self._link_hover(lbl, font, True))
        lbl.bind("<Leave>", lambda _e: self._link_hover(lbl, font, False))
        lbl.bind("<Button-1>", lambda _e: self._link_click(lbl, command))
        self._links.append((lbl, font, command))
        return lbl

    def _link_hover(self, lbl, font, on):
        if lbl.cget("fg") == FAINT:
            return
        font.configure(underline=on)

    def _link_click(self, lbl, command):
        if lbl.cget("fg") != FAINT:
            command()

    def _build_toolbar(self):
        bar = tk.Frame(self.root, bg=BG)
        bar.pack(fill=tk.X, padx=14, pady=(10, 8))
        self._make_link(bar, "添加文件", self.add_files)
        self._make_link(bar, "添加文件夹", self.add_folder)
        self._make_link(bar, "清空", self.clear_files)
        self.tools_link = self._make_link(bar, "PDF 工具", self._toggle_pdf_menu)
        self.count_lbl = tk.Label(bar, text="0 项", font=self._f["count"],
                                  bg=BG, fg=FAINT)
        self.count_lbl.pack(side=tk.RIGHT)

    def _build_dropzone(self):
        self.drop_canvas = tk.Canvas(self.root, bg=PANEL, highlightthickness=0, bd=0)
        self.drop_canvas.pack(fill=tk.BOTH, expand=True, padx=14)
        self.drop_canvas.bind("<Configure>", self._draw_dropzone)

        self.drop_content = tk.Frame(self.drop_canvas, bg=PANEL)
        self._drop_window = self.drop_canvas.create_window(
            2, 2, anchor=tk.NW, window=self.drop_content)

        # 空状态
        self.empty = tk.Frame(self.drop_content, bg=PANEL, cursor="hand2")
        icon = tk.Canvas(self.empty, width=26, height=26, bg=PANEL,
                         highlightthickness=0)
        icon.pack(pady=(0, 6))
        c = LINE
        icon.create_line(13, 21, 13, 5, fill=c, width=1)
        icon.create_line(13, 5, 8, 10, fill=c, width=1)
        icon.create_line(13, 5, 18, 10, fill=c, width=1)
        icon.create_line(4, 19, 4, 23, 22, 23, 22, 19, fill=c, width=1)
        t1 = tk.Label(self.empty, text="拖入文件开始", font=self._f["empty"],
                      bg=PANEL, fg=MUTED)
        t1.pack()
        t2 = tk.Label(self.empty, text="或点击选择", font=self._f["meta"],
                      bg=PANEL, fg=FAINT)
        t2.pack()
        for w in (self.empty, icon, t1, t2):
            w.bind("<Button-1>", lambda _e: self._empty_click())

        # 文件行列表
        self.rows_canvas = tk.Canvas(self.drop_content, bg=PANEL,
                                     highlightthickness=0)
        self.rows_inner = tk.Frame(self.rows_canvas, bg=PANEL)
        self._rows_window = self.rows_canvas.create_window(
            0, 0, anchor=tk.NW, window=self.rows_inner)
        self.rows_scroll = ttk.Scrollbar(self.drop_content, style="Slim.Vertical.TScrollbar",
                                         orient=tk.VERTICAL, command=self.rows_canvas.yview)
        self.rows_canvas.configure(yscrollcommand=self.rows_scroll.set)
        self.rows_inner.bind("<Configure>", self._on_rows_configure)
        self.rows_canvas.bind("<Configure>", self._on_rows_canvas_configure)
        self.rows_canvas.bind("<MouseWheel>", self._on_wheel)

        if HAS_DND:
            self._register_dnd()

    def _register_dnd(self):
        for w in (self.drop_canvas, self.drop_content, self.empty,
                  self.rows_canvas, self.rows_inner):
            try:
                w.drop_target_register(tkinterdnd2.DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_drop)
                w.dnd_bind("<<DragEnter>>", lambda _e: self._set_drag(True))
                w.dnd_bind("<<DragLeave>>", lambda _e: self._set_drag(False))
            except tk.TclError:
                pass

    def _build_footer(self):
        foot = tk.Frame(self.root, bg=BG)
        foot.pack(side=tk.BOTTOM, fill=tk.X, padx=14, pady=(12, 14))

        main = tk.Frame(foot, bg=BG)
        main.pack(fill=tk.X)
        fmt = tk.Frame(main, bg=BG)
        fmt.pack(side=tk.LEFT)
        tk.Label(fmt, text="目标格式", font=self._f["label"], bg=BG,
                 fg=FAINT).pack(anchor=tk.W)
        self.target = ttk.Combobox(fmt, width=12, state="readonly",
                                   style="Flat.TCombobox", font=self._f["link"])
        self.target.pack()
        tk.Frame(fmt, bg=LINE, height=1).pack(fill=tk.X)
        self.more_link = self._make_link(main, "选项 ▾", self._toggle_advanced,
                                         side=tk.LEFT, padx=(12, 0))

        # 折叠选项
        self.adv = tk.Frame(foot, bg=BG)
        out = tk.Frame(self.adv, bg=BG)
        out.pack(side=tk.LEFT)
        tk.Label(out, text="输出目录", font=self._f["label"], bg=BG,
                 fg=FAINT).pack(anchor=tk.W)
        out_row = tk.Frame(out, bg=BG)
        out_row.pack()
        self.outdir = tk.Entry(out_row, width=20, relief="flat", bg=BG,
                               fg=MUTED, font=self._f["meta"])
        self.outdir.pack(side=tk.LEFT)
        browse_font = tkfont.Font(root=self.root, family=FONT, size=-10)
        browse = tk.Label(out_row, text="浏览", font=browse_font, bg=BG,
                          fg=MUTED, cursor="hand2")
        browse.pack(side=tk.LEFT, padx=(6, 0))
        browse.bind("<Enter>", lambda _e: browse_font.configure(underline=True))
        browse.bind("<Leave>", lambda _e: browse_font.configure(underline=False))
        browse.bind("<Button-1>", lambda _e: self.pick_outdir())
        tk.Frame(out, bg=LINE, height=1).pack(fill=tk.X)

        dpi_f = tk.Frame(self.adv, bg=BG)
        dpi_f.pack(side=tk.LEFT, padx=(16, 0))
        tk.Label(dpi_f, text="DPI", font=self._f["label"], bg=BG,
                 fg=FAINT).pack(anchor=tk.W)
        self.dpi = tk.Entry(dpi_f, width=5, relief="flat", bg=BG, fg=TEXT,
                            font=self._f["mono"])
        self.dpi.insert(0, "150")
        self.dpi.pack()
        self.dpi.bind("<FocusOut>", lambda _e: self._normalize_dpi())
        tk.Frame(dpi_f, bg=LINE, height=1).pack(fill=tk.X)

        enc_f = tk.Frame(self.adv, bg=BG)
        enc_f.pack(side=tk.LEFT, padx=(16, 0))
        tk.Label(enc_f, text="编码", font=self._f["label"], bg=BG,
                 fg=FAINT).pack(anchor=tk.W)
        self.encoding = ttk.Combobox(enc_f, width=7, state="readonly",
                                     style="Flat.TCombobox", font=self._f["meta"],
                                     values=ENCODINGS)
        self.encoding.set("utf-8")
        self.encoding.pack()
        tk.Frame(enc_f, bg=LINE, height=1).pack(fill=tk.X)

        # 运行行（固定宽度控件先 pack，弹性进度条最后）
        run = tk.Frame(foot, bg=BG)
        run.pack(fill=tk.X, pady=(12, 0))
        self.cta = tk.Canvas(run, width=104, height=30, bg=BG,
                             highlightthickness=0, cursor="hand2")
        self.cta.pack(side=tk.RIGHT)
        self.status_lbl = tk.Label(run, text="就绪", font=self._f["meta"],
                                   bg=BG, fg=FAINT, anchor=tk.W, width=24)
        self.status_lbl.pack(side=tk.RIGHT, padx=(0, 12), pady=(6, 0))
        self.track = tk.Canvas(run, height=2, bg=LINE_SOFT,
                               highlightthickness=0, bd=0)
        self.track.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=(14, 0))
        self._fill = self.track.create_rectangle(0, 0, 0, 2, fill=ACCENT, width=0)
        self._cta_hover = False
        self.cta.bind("<Enter>", lambda _e: self._cta_set_hover(True))
        self.cta.bind("<Leave>", lambda _e: self._cta_set_hover(False))
        self.cta.bind("<Button-1>", lambda _e: self._cta_click())
        self._draw_cta()

    # ---- 外观状态 ----
    def _draw_dropzone(self, _e=None):
        c = self.drop_canvas
        c.delete("border")
        w, h = c.winfo_width(), c.winfo_height()
        if w < 4 or h < 4:
            return
        c.coords(self._drop_window, 2, 2)
        c.itemconfigure(self._drop_window, width=w - 4, height=h - 4)
        if self._drag_over:
            outline, dash = ACCENT, ()
        elif self.files:
            outline, dash = LINE_SOFT, ()
        else:
            outline, dash = LINE, (4, 3)
        c.create_rectangle(1, 1, w - 1, h - 1, outline=outline, dash=dash,
                           tags="border")

    def _set_drag(self, on):
        if self.running:
            return
        if on != self._drag_over:
            self._drag_over = on
            self._draw_dropzone()

    def _cta_set_hover(self, on):
        self._cta_hover = on
        self._draw_cta()

    def _cta_click(self):
        if not self.running and self.files:
            self.start_convert()

    def _draw_cta(self):
        c = self.cta
        c.delete("all")
        w, h = 104, 30
        if self.running:
            text, outline, fill, fg = "转换中…", ACCENT, "", ACCENT_DEEP
            c.configure(cursor="arrow")
        elif not self.files:
            text, outline, fill, fg = "开始转换", LINE, "", FAINT
            c.configure(cursor="arrow")
        elif self._cta_hover:
            text, outline, fill, fg = "开始转换", TEXT, TEXT, BG
            c.configure(cursor="hand2")
        else:
            text, outline, fill, fg = "开始转换", TEXT, "", TEXT
            c.configure(cursor="hand2")
        _round_rect(c, 1, 1, w - 1, h - 1, h // 2 - 1,
                    outline=outline, fill=fill, width=1)
        c.create_text(w / 2, h / 2, text=text, fill=fg, font=self._f["link"])

    def _set_status(self, text):
        self.status_lbl.configure(text=text)

    def _set_running(self, on):
        self.running = on
        for lbl, _font, _cmd in self._links:
            lbl.configure(fg=FAINT if on else TEXT, cursor="arrow" if on else "hand2")
        self._draw_cta()

    def _toggle_advanced(self):
        self._adv_open = not self._adv_open
        self.more_link.configure(text="选项 ▴" if self._adv_open else "选项 ▾")
        if self._adv_open:
            self.adv.pack(fill=tk.X, pady=(10, 2),
                          before=self.status_lbl.master)
        else:
            self.adv.pack_forget()

    def _toggle_pdf_menu(self):
        menu = tk.Menu(self.root, tearoff=0, bg=PANEL, fg=TEXT,
                       activebackground=MENU_HOVER, activeforeground=TEXT,
                       disabledforeground=FAINT, bd=1, relief=tk.SOLID,
                       font=self._f["link"])
        menu.add_command(label="合并 PDF", command=self.merge_pdf)
        menu.add_command(label="拆分 PDF", command=self.split_pdf)
        menu.add_command(label="旋转 PDF", command=self.rotate_pdf)
        menu.add_separator()
        menu.add_command(label="作用于列表中的 PDF", state=tk.DISABLED)
        x = self.tools_link.winfo_rootx()
        y = self.tools_link.winfo_rooty() + self.tools_link.winfo_height() + 2
        try:
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    # ---- 文件行 ----
    def _on_rows_configure(self, _e=None):
        self.rows_canvas.configure(
            scrollregion=self.rows_canvas.bbox("all"))
        self.rows_scroll.pack_forget()
        if self.rows_inner.winfo_reqheight() > self.rows_canvas.winfo_height():
            self.rows_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    def _on_rows_canvas_configure(self, e):
        self.rows_canvas.itemconfigure(self._rows_window, width=e.width)
        self._render_rows()

    def _on_wheel(self, e):
        self.rows_canvas.yview_scroll(int(-e.delta / 120), "units")

    def _bind_row(self, w, idx):
        w.bind("<Enter>", lambda _e: self._row_hover(idx, True))
        w.bind("<Leave>", lambda _e: self._row_hover(idx, False))
        w.bind("<MouseWheel>", self._on_wheel)

    def _row_hover(self, idx, on):
        if idx >= len(self._row_widgets):
            return
        widgets = self._row_widgets[idx]
        bg = ROW_HOVER if on else PANEL
        for w in widgets["paint"]:
            try:
                w.configure(bg=bg)
            except tk.TclError:
                pass
        widgets["rm"].configure(fg=FAINT if on else (ROW_HOVER if on else PANEL))

    def _render_rows(self):
        for child in self.rows_inner.winfo_children():
            child.destroy()
        self._row_widgets = []
        try:
            name_px = max(
                40,
                self.rows_canvas.winfo_width() - (20 + 34 + 62 + 58 + 22 + 32))
        except tk.TclError:
            name_px = 200
        for i, f in enumerate(self.files):
            self._build_row(i, f, name_px)
        # 空状态 / 列表切换
        if self.files:
            self.empty.pack_forget()
            if not self.rows_canvas.winfo_ismapped():
                self.rows_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        else:
            self.rows_canvas.pack_forget()
            self.rows_scroll.pack_forget()
            self.empty.pack(fill=tk.BOTH, expand=True)
        n = len(self.files)
        self.count_lbl.configure(text=f"{n} 项")
        self._draw_dropzone()
        self._draw_cta()

    def _build_row(self, idx, f, name_px):
        row = tk.Frame(self.rows_inner, bg=PANEL, height=38)
        row.pack(fill=tk.X)
        row.pack_propagate(False)
        paint = [row]

        badge = tk.Frame(row, bg=PANEL, width=28, height=20,
                         highlightthickness=1, highlightbackground=LINE)
        badge.pack_propagate(False)
        ext = f["path"].suffix.lstrip(".").upper()[:4] or "FILE"
        bl = tk.Label(badge, text=ext, font=self._f["ext"], bg=PANEL, fg=MUTED)
        bl.pack(expand=True)
        badge.grid(row=0, column=0, padx=(10, 8))
        paint += [badge, bl]

        name = self._ellipsize(f["path"].name, self._f["name"], name_px)
        nl = tk.Label(row, text=name, font=self._f["name"], bg=PANEL, fg=TEXT,
                      anchor=tk.W)
        nl.grid(row=0, column=1, sticky=tk.EW)
        paint.append(nl)

        try:
            size = _fmt_size(f["path"].stat().st_size)
        except OSError:
            size = "—"
        sl = tk.Label(row, text=size, font=self._f["mono"], bg=PANEL, fg=FAINT,
                      anchor=tk.E)
        sl.grid(row=0, column=2, padx=(8, 8), sticky=tk.E)
        paint.append(sl)

        st = tk.Frame(row, bg=PANEL)
        st.grid(row=0, column=3, padx=(0, 8), sticky=tk.E)
        dot = tk.Canvas(st, width=7, height=7, bg=PANEL, highlightthickness=0)
        dot.pack(side=tk.LEFT, padx=(0, 5))
        tl = tk.Label(st, font=self._f["meta"], bg=PANEL)
        tl.pack(side=tk.LEFT)
        paint += [st, dot, tl]
        self._paint_status(dot, tl, f["state"])

        rm = tk.Label(row, text="×", font=self._f["name"], bg=PANEL, fg=PANEL,
                      cursor="hand2")
        rm.grid(row=0, column=4, padx=(0, 10))
        rm.bind("<Button-1>", lambda _e, i=idx: self._remove_at(i))
        paint.append(rm)

        row.grid_columnconfigure(1, weight=1)
        for col, wpx in ((0, 28), (2, 62), (3, 58), (4, 14)):
            row.grid_columnconfigure(col, minsize=wpx)
        self._row_widgets.append({"paint": paint, "rm": rm, "dot": dot, "text": tl})
        for w in paint:
            self._bind_row(w, idx)

    def _paint_status(self, dot, tl, state):
        dot.delete("all")
        if state == "doing":
            dot.create_oval(1, 1, 6, 6, fill=ACCENT, outline=ACCENT)
            tl.configure(text=STATE_TEXT[state], fg=ACCENT_DEEP)
        elif state == "done":
            dot.create_oval(1, 1, 6, 6, fill=TEXT, outline=TEXT)
            tl.configure(text=STATE_TEXT[state], fg=MUTED)
        elif state == "fail":
            dot.create_oval(1, 1, 6, 6, fill=TEXT, outline=TEXT)
            tl.configure(text=STATE_TEXT[state], fg=TEXT)
        else:
            dot.create_oval(1, 1, 6, 6, fill="", outline=LINE)
            tl.configure(text=STATE_TEXT[state], fg=FAINT)

    def _set_row_state(self, idx, state):
        if idx >= len(self.files):
            return
        self.files[idx]["state"] = state
        w = self._row_widgets[idx]
        self._paint_status(w["dot"], w["text"], state)

    def _ellipsize(self, text, font, max_px):
        if font.measure(text) <= max_px:
            return text
        while text and font.measure(text + "…") > max_px:
            text = text[:-1]
        return text + "…" if text else "…"

    # ---- 文件列表 ----
    def add_files(self):
        for p in filedialog.askopenfilenames(title="选择要转换的文件"):
            self._add(Path(p))

    def add_folder(self):
        folder = filedialog.askdirectory(title="选择文件夹")
        if folder:
            for p in sorted(Path(folder).rglob("*")):
                if p.is_file():
                    self._add(p)

    def _add(self, p: Path):
        p = Path(p)
        if any(f["path"] == p for f in self.files):
            return
        self.files.append({"path": p, "state": "idle", "err": None})
        self.refresh_targets()
        self._render_rows()
        self._set_status(f"添加 {len(self.files)} 项 · 目标 {self.target.get().upper() or '—'}")

    def _remove_at(self, idx):
        if self.running or idx >= len(self.files):
            return
        name = self.files.pop(idx)["path"].name
        self.refresh_targets()
        self._render_rows()
        self._set_status(f"已移除 {name}")

    def clear_files(self):
        self.files.clear()
        self.track.coords(self._fill, 0, 0, 0, 2)
        self.refresh_targets()
        self._render_rows()
        self._set_status("已清空")

    def _empty_click(self):
        if not self.running:
            self.add_files()

    def _on_drop(self, event):
        self._set_drag(False)
        if self.running:
            return
        try:
            paths = self.root.tk.splitlist(event.data)
        except tk.TclError:
            paths = (event.data,)
        for raw in paths:
            p = Path(raw)
            if p.is_dir():
                for q in sorted(p.rglob("*")):
                    if q.is_file():
                        self._add(q)
            elif p.is_file():
                self._add(p)

    def refresh_targets(self):
        exts = {f["path"].suffix.lower().lstrip(".") for f in self.files}
        common: set[str] | None = None
        for ext in exts:
            targets = set(registry.available_targets(ext))
            common = targets if common is None else common & targets
        values = sorted(common or [])
        self.target.configure(values=values)
        if not values:
            self.target.set("")
        elif self.target.get() not in values:
            self.target.set(values[0])

    def pick_outdir(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.outdir.delete(0, tk.END)
            self.outdir.insert(0, d)

    def _dpi_value(self) -> int:
        try:
            v = int(self.dpi.get().strip() or 150)
        except ValueError:
            v = 150
        return max(72, min(600, v))

    def _normalize_dpi(self):
        self.dpi.delete(0, tk.END)
        self.dpi.insert(0, str(self._dpi_value()))

    # ---- 转换 ----
    def start_convert(self):
        if self.running or not self.files:
            if not self.files:
                self._set_status("请先添加要转换的文件")
            return
        target = self.target.get()
        if not target:
            self._set_status("请选择目标格式")
            return
        outdir = self.outdir.get().strip() or None
        dpi = self._dpi_value()
        dst_encoding = self.encoding.get() or "utf-8"
        items = [(i, f["path"]) for i, f in enumerate(self.files)
                 if f["state"] != "done"]
        if not items:
            self._set_status("全部已转换过")
            return
        self._set_running(True)
        self.track.coords(self._fill, 0, 0, 0, 2)
        threading.Thread(
            target=self._worker,
            args=(items, target, outdir, dpi, dst_encoding),
            daemon=True,
        ).start()
        self.root.after(100, self._poll)

    def _worker(self, items, target, outdir, dpi, dst_encoding):
        ok, fail = 0, 0
        total = len(items)
        for n, (idx, src) in enumerate(items):
            self.queue.put(("doing", (idx, n, total, src.name)))
            try:
                dst = None
                if outdir:
                    dst = registry.unique_path(Path(outdir) / (src.stem + "." + target))
                registry.convert(src, target, dst, dpi=dpi, dst_encoding=dst_encoding)
                self.queue.put(("row", (idx, "done", src.name, None)))
                ok += 1
            except ConvertError as e:
                self.queue.put(("row", (idx, "fail", src.name, str(e))))
                fail += 1
            self.queue.put(("step", (ok + fail, total)))
        self.queue.put(("done", (ok, fail, outdir)))

    def _poll(self):
        if not self._drain():
            self.root.after(100, self._poll)

    def _drain(self) -> bool:
        """处理队列；返回 True 表示本批转换已全部结束。"""
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "doing":
                    idx, n, total, name = payload
                    self._set_row_state(idx, "doing")
                    self._set_status(f"{n + 1}/{total} {name}")
                elif kind == "row":
                    idx, state, name, err = payload
                    self._set_row_state(idx, state)
                    if state == "fail":
                        self._set_status(f"失败：{name} — {err}")
                elif kind == "step":
                    done, total = payload
                    w = max(self.track.winfo_width(), 1)
                    self.track.coords(self._fill, 0, 0, w * done / total, 2)
                elif kind == "done":
                    ok, fail, outdir = payload
                    self._set_running(False)
                    w = max(self.track.winfo_width(), 1)
                    self.track.coords(self._fill, 0, 0, w, 2)
                    dest = outdir or "源文件目录"
                    self._set_status(f"完成 {ok}/{ok + fail} · 输出至 {dest}")
                    return True
        except queue.Empty:
            pass
        return False

    # ---- PDF 工具 ----
    def _listed_pdfs(self) -> list[Path]:
        return [f["path"] for f in self.files
                if f["path"].suffix.lower() == ".pdf"]

    def merge_pdf(self):
        if self.running:
            return
        pdfs = self._listed_pdfs()
        if len(pdfs) < 2:
            self._set_status("合并需要至少 2 个 PDF")
            return
        out = filedialog.asksaveasfilename(defaultextension=".pdf",
                                           filetypes=[("PDF", "*.pdf")])
        if not out:
            return
        try:
            merge_pdfs(pdfs, Path(out))
            self._set_status(f"已合并 {len(pdfs)} 个 PDF → {out}")
        except ConvertError as e:
            self._set_status(f"失败：{e}")

    def split_pdf(self):
        if self.running:
            return
        pdfs = self._listed_pdfs()
        if len(pdfs) != 1:
            self._set_status("拆分需要列表中恰好 1 个 PDF")
            return
        pages = simpledialog.askstring("拆分 PDF", "页码范围（如 1-3,5）：")
        if not pages:
            return
        outdir = filedialog.askdirectory(title="选择输出目录")
        if not outdir:
            return
        try:
            outs = split_pdf(pdfs[0], Path(outdir), pages)
            self._set_status("已拆分出 " + ", ".join(o.name for o in outs))
        except ConvertError as e:
            self._set_status(f"失败：{e}")

    def rotate_pdf(self):
        if self.running:
            return
        pdfs = self._listed_pdfs()
        if len(pdfs) != 1:
            self._set_status("旋转需要列表中恰好 1 个 PDF")
            return
        angle = simpledialog.askinteger("旋转 PDF", "角度（90/180/270）：",
                                        initialvalue=90)
        if angle not in (90, 180, 270):
            return
        out = filedialog.asksaveasfilename(defaultextension=".pdf",
                                           filetypes=[("PDF", "*.pdf")])
        if not out:
            return
        try:
            rotate_pdf(pdfs[0], Path(out), angle)
            self._set_status(f"已旋转 {angle}° → {out}")
        except ConvertError as e:
            self._set_status(f"失败：{e}")


def run():
    if HAS_DND:
        root = tkinterdnd2.TkinterDnD.Tk()
    else:
        root = tk.Tk()
    ConverterApp(root)
    root.mainloop()
