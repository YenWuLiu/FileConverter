import os
import sys

if getattr(sys, "frozen", False):
    _DIR = os.path.dirname(sys.executable)
    _LOG = os.path.join(_DIR, "boot_gui.log")
    import faulthandler

    _fh = open(os.path.join(_DIR, "gui_fault.log"), "w", encoding="utf-8")

    faulthandler.enable(file=_fh)
    # pyi_rth__tkinter 设置 TCL_LIBRARY/TK_LIBRARY 环境变量，但 Tk 8.6 的 C 层
    # 在进程首次初始化时会缓存库路径（本构建中缓存的是 exe 目录而非 _internal），
    # 导致后续 Tk() 仍报 "Can't find a usable init.tcl"。这里在创建任何 Tk
    # 解释器之前通过 _tkinter 直接修正为打包内的 _tcl_data/_tk_data。
    try:
        import _tkinter

        _base = sys._MEIPASS
        os.environ["TCL_LIBRARY"] = os.path.join(_base, "_tcl_data")
        os.environ["TK_LIBRARY"] = os.path.join(_base, "_tk_data")
    except Exception:
        pass
else:
    _LOG = os.devnull


def _t(msg):
    try:
        with open(_LOG, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except OSError:
        pass


try:
    _t("0 entry")
    import tkinter as tk
    import tkinterdnd2

    _t(f"1 imports ok; TkVersion={tk.TkVersion}")
    try:
        root = tkinterdnd2.TkinterDnD.Tk()
        _t("2 TkinterDnD.Tk ok")
    except tk.TclError as e:
        _t(f"2 TkinterDnD.Tk TclError: {e!r}")
        root = tk.Tk()
        _t("2 plain tk.Tk ok")
    _t(f"3 tcl patchlevel={root.tk.call('info', 'patchlevel')}")
    try:
        _t(f"3 tkdnd present={root.tk.call('package', 'present', 'tkdnd')}")
    except tk.TclError as e:
        _t(f"3 tkdnd present check TclError: {e!r}")
    from file_converter.gui import ConverterApp

    ConverterApp(root)
    _t(f"4 app built; exists={root.winfo_exists()}")
    root.update()
    _t(f"5 after update: exists={root.winfo_exists()} mapped={root.winfo_ismapped()} geometry={root.winfo_geometry()}")
    root.mainloop()
    _t("6 mainloop returned")
except BaseException:
    import traceback

    with open(_LOG, "a", encoding="utf-8") as f:
        traceback.print_exc(file=f)
    raise
