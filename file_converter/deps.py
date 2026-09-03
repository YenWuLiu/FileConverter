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


def office_available() -> bool:
    """探测本机 MS Office COM 是否可用（以 Word 为探针）。"""
    if not module_available("win32com.client"):
        return False
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        try:
            app = win32com.client.DispatchEx("Word.Application")
            app.Quit()
            return True
        finally:
            pythoncom.CoUninitialize()
    except Exception:
        return False


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
