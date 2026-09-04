# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec：onedir 单文件夹双 exe。

- FileConverter.exe    窗口程序（无控制台），启动 Tkinter GUI
- FileConverterCLI.exe 控制台程序，运行 CLI（file_converter.__main__:main）

构建（在项目根目录执行，相对路径以根目录为准）：
    pyinstaller packaging/FileConverter.spec --noconfirm --clean
输出：dist/FileConverter/ —— 两个 exe 与 _internal 依赖目录必须整体分发。
"""
import os

from PyInstaller.utils.hooks import collect_data_files

SPEC_DIR = os.path.abspath(SPECPATH)  # packaging/（PyInstaller 注入的 spec 所在目录）
ROOT = os.path.dirname(SPEC_DIR)  # 项目根（file_converter 包所在）

# imageio_ffmpeg 的 ffmpeg exe（83MB）必须作为真实文件落盘到
# _internal/imageio_ffmpeg/binaries/：运行时 imageio_ffmpeg.get_ffmpeg_exe()
# 通过 importlib.resources 定位该目录并直接执行其中的 exe。
# pyinstaller-hooks-contrib 的 hook-imageio_ffmpeg 也会做同样的收集，这里显式兜底。
ffmpeg_datas = collect_data_files("imageio_ffmpeg", subdir="binaries")

# tkinterdnd2 的 tkdnd 原生二进制（libtkdnd*.dll 等，GUI 拖放）。
# contrib hook-tkinterdnd2 也会收集，这里显式收集兜底（重复条目会被去重）。
tkdnd_datas = collect_data_files("tkinterdnd2")

# GUI 图标：冻结后运行时通过 sys._MEIPASS/file_converter/assets/ 访问
assets_datas = [(os.path.join(ROOT, "file_converter", "assets"), "file_converter/assets")]

a = Analysis(
    [os.path.join(SPEC_DIR, "entry_gui.py"), os.path.join(SPEC_DIR, "entry_cli.py")],
    pathex=[ROOT],
    binaries=[],
    datas=ffmpeg_datas + tkdnd_datas + assets_datas,
    hiddenimports=[
        # COM（docx/xlsx/pptx ↔ pdf）在函数体内 lazy import，显式声明保险
        "win32com.client",
        "pythoncom",
        "pywintypes",
        # get_ffmpeg_exe() 用 importlib.resources 访问该包
        "imageio_ffmpeg.binaries",
        # 其余第三方依赖（fitz/pymupdf、cv2、numpy、fire、fonttools、markdown、
        # bs4、yaml、docx、pptx、openpyxl、pypdf、PIL）均由 modulegraph 跟随
        # 源码中的 import 自动收集，数据文件由对应 hook（docx/pptx/openpyxl/
        # cv2/markdown/PIL/numpy 等）处理。
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

# 入口脚本 TOC 条目（name, path, typecode）。不要对 a.scripts 切片：它是已弃用的
# TOC 对象，切片会经 list 基类返回普通 list[str]，被 EXE 当作 TOC 静默逐字符
# 展开，导致入口脚本丢失（构建不报错，但 exe 运行的是错误入口）。
def _script_entry(script_name):
    for dest, src, typecode in a.scripts:
        if dest == script_name:
            return [(dest, src, typecode)]
    raise SystemExit(f"脚本 {script_name} 不在 a.scripts 中")

exe_gui = EXE(
    pyz,
    _script_entry("entry_gui"),
    exclude_binaries=True,
    name="FileConverter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)
exe_cli = EXE(
    pyz,
    _script_entry("entry_cli"),
    exclude_binaries=True,
    name="FileConverterCLI",
    debug=True,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
)
coll = COLLECT(
    exe_gui,
    exe_cli,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="FileConverter",
)
