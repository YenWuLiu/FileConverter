import importlib.util
import os
import shutil

_cache: dict[str, bool] = {}


def module_available(name: str) -> bool:
    if name not in _cache:
        try:
            _cache[name] = importlib.util.find_spec(name) is not None
        except Exception:
            _cache[name] = False
    return _cache[name]


_OFFICE_AVAILABLE: bool | None = None


def office_available() -> bool:
    """探测本机 MS Office COM 是否可用（以 Word 为探针，每进程最多探测一次）。"""
    global _OFFICE_AVAILABLE
    if _OFFICE_AVAILABLE is not None:
        return _OFFICE_AVAILABLE
    if not module_available("win32com.client"):
        _OFFICE_AVAILABLE = False
        return False
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        try:
            app = win32com.client.DispatchEx("Word.Application")
            app.Quit()
            _OFFICE_AVAILABLE = True
        finally:
            pythoncom.CoUninitialize()
    except Exception:
        _OFFICE_AVAILABLE = False
    return _OFFICE_AVAILABLE


def libreoffice_path() -> str | None:
    exe = shutil.which("soffice")
    if exe:
        return exe
    for env in ("ProgramFiles", "ProgramFiles(x86)"):
        root = os.environ.get(env)
        if root:
            p = os.path.join(root, "LibreOffice", "program", "soffice.exe")
            if os.path.exists(p):
                return p
    return None
