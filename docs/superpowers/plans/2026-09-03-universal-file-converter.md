# 万能文件转换工具 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 构建一个 Windows 桌面万能文件转换工具（GUI + CLI），覆盖日常办公文件类型转换。

**Architecture:** 转换器注册表模式：`converters/base.py` 提供 `@register` 装饰器与 `ConvertError`；各 converter 模块注册 `(源格式集, 目标格式集) → 函数`；`registry.py` 统一查找/调用并做依赖过滤；`deps.py` 探测可选依赖（PIL/docx/openpyxl/pptx/pypdf/fitz/win32com/LibreOffice）；CLI 与 Tkinter GUI 是两个薄前端。

**Tech Stack:** Python 3.14；Pillow、python-docx、openpyxl、python-pptx、pypdf、PyMuPDF、pywin32、markdown、beautifulsoup4、PyYAML；pytest。

**Spec:** `docs/superpowers/specs/2026-09-03-universal-file-converter-design.md`

## Global Constraints

- 工作目录即项目根：`D:\Office Tools\File Converter`
- Python 解释器：`C:\Python314\python.exe`（命令行用 `python`）
- 所有用户可见的错误信息用中文；批量转换单文件失败不得中断，结束给汇总
- 输出文件重名时自动加 ` (1)` 后缀，绝不覆盖已有文件（`registry.unique_path`）
- 依赖缺失的转换不得出现在可选目标中；调用时报清晰中文错误
- 测试只覆盖纯库路径，不测 Office COM / LibreOffice / GUI
- 编码约定：CSV 写出用 `utf-8-sig`（Excel 兼容），文本默认 UTF-8

---

### Task 1: 项目骨架（registry / deps / base / requirements / git）

**Files:**
- Create: `requirements.txt`
- Create: `file_converter/__init__.py`
- Create: `file_converter/converters/__init__.py`
- Create: `file_converter/converters/base.py`
- Create: `file_converter/deps.py`
- Create: `file_converter/registry.py`
- Test: `tests/test_registry.py`

**Interfaces:**
- Produces:
  - `base.register(src, dst, requires=(), label="")` 装饰器；`base.REGISTRY: list[ConverterSpec]`；`base.ConvertError`
  - `registry.available_targets(src_ext: str) -> list[str]`
  - `registry.convert(src_path, dst_ext: str, dst_path=None, **opts) -> Path`
  - `registry.unique_path(p: Path) -> Path`
  - `registry.matrix() -> list[tuple[tuple[str,...], tuple[str,...], str, bool]]`（src, dst, label, 依赖可用）
  - `deps.module_available(name: str) -> bool`、`deps.office_available() -> bool`、`deps.libreoffice_path() -> str | None`

- [ ] **Step 1: git init + requirements.txt**

```bash
cd "D:\Office Tools\File Converter" && git init
```

`requirements.txt`：
```
Pillow
python-docx
openpyxl
python-pptx
pypdf
PyMuPDF
pywin32; sys_platform == "win32"
markdown
beautifulsoup4
PyYAML
pytest
```

安装：`python -m pip install -r requirements.txt`（超时给 300s）

- [ ] **Step 2: 写失败测试 `tests/test_registry.py`**

```python
from pathlib import Path

import pytest

from file_converter import registry
from file_converter.converters.base import ConvertError, register


def test_register_and_available_targets():
    @register("foo", ["bar", "baz"], label="测试")
    def _foo(src, dst, **opts):
        dst.write_text("ok", encoding="utf-8")

    assert registry.available_targets("foo") == ["bar", "baz"]


def test_convert_writes_unique_output(tmp_path):
    @register("aaa", ["bbb"])
    def _conv(src, dst, **opts):
        dst.write_text("x", encoding="utf-8")

    src = tmp_path / "f.aaa"
    src.write_text("1", encoding="utf-8")
    out1 = registry.convert(src, "bbb")
    out2 = registry.convert(src, "bbb")
    assert out1.name == "f.bbb"
    assert out2.name == "f (1).bbb"


def test_convert_unsupported_raises(tmp_path):
    src = tmp_path / "f.zzz"
    src.write_text("1", encoding="utf-8")
    with pytest.raises(ConvertError, match="不支持"):
        registry.convert(src, "qqq")


def test_convert_wraps_errors(tmp_path):
    @register("err", ["out"])
    def _boom(src, dst, **opts):
        raise ValueError("底层炸了")

    src = tmp_path / "f.err"
    src.write_text("1", encoding="utf-8")
    with pytest.raises(ConvertError, match="转换失败"):
        registry.convert(src, "out")


def test_missing_dependency_filtered():
    @register("dep", ["x"], requires=("nonexistent_module_xyz",))
    def _d(src, dst, **opts):
        pass

    assert "x" not in registry.available_targets("dep")
    src = Path("whatever.dep")
    with pytest.raises(ConvertError, match="缺少依赖"):
        registry.convert(src, "x")
```

- [ ] **Step 3: 运行测试确认失败**

Run: `python -m pytest tests/test_registry.py -v`
Expected: FAIL（`ModuleNotFoundError: file_converter`）

- [ ] **Step 4: 实现骨架**

`file_converter/__init__.py`：
```python
__version__ = "1.0.0"
```

`file_converter/converters/base.py`：
```python
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable


class ConvertError(Exception):
    """单个文件转换失败时抛出，message 面向最终用户（中文）。"""


ConvertFunc = Callable[..., None]  # (src: Path, dst: Path, **opts) -> None


@dataclass
class ConverterSpec:
    src: tuple[str, ...]
    dst: tuple[str, ...]
    func: ConvertFunc
    requires: tuple[str, ...] = ()
    label: str = ""


REGISTRY: list[ConverterSpec] = []


def _norm(values) -> tuple[str, ...]:
    if isinstance(values, str):
        values = [values]
    return tuple(v.lower().lstrip(".") for v in values)


def register(src, dst, requires=(), label=""):
    """注册一个转换器：src/dst 为扩展名（可单个或列表），requires 为必需的 import 模块名。"""

    def deco(fn: ConvertFunc) -> ConvertFunc:
        REGISTRY.append(
            ConverterSpec(_norm(src), _norm(dst), fn, tuple(requires), label or fn.__name__)
        )
        return fn

    return deco
```

`file_converter/converters/__init__.py`：
```python
from . import base, documents, images, pdf, textdata  # noqa: F401  导入即注册
```

`file_converter/deps.py`：
```python
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
```

`file_converter/registry.py`：
```python
from __future__ import annotations

from pathlib import Path

from . import converters  # noqa: F401  确保所有 converter 已注册
from . import deps
from .converters.base import REGISTRY, ConvertError


def _spec_ok(spec) -> bool:
    return all(deps.module_available(m) for m in spec.requires)


def available_targets(src_ext: str) -> list[str]:
    src_ext = src_ext.lower().lstrip(".")
    out: set[str] = set()
    for spec in REGISTRY:
        if src_ext in spec.src and _spec_ok(spec):
            for d in spec.dst:
                if d != src_ext or spec.label:
                    out.add(d)
    return sorted(out)


def matrix() -> list[tuple[tuple[str, ...], tuple[str, ...], str, bool]]:
    return [(s.src, s.dst, s.label, _spec_ok(s)) for s in REGISTRY]


def unique_path(p: Path) -> Path:
    p = Path(p)
    if not p.exists():
        return p
    for i in range(1, 1000):
        q = p.with_name(f"{p.stem} ({i}){p.suffix}")
        if not q.exists():
            return q
    raise ConvertError(f"无法为 {p.name} 生成不重名的输出文件")


def convert(src_path, dst_ext: str, dst_path=None, **opts) -> Path:
    src = Path(src_path)
    dst_ext = dst_ext.lower().lstrip(".")
    src_ext = src.suffix.lower().lstrip(".")
    for spec in REGISTRY:
        if src_ext in spec.src and dst_ext in spec.dst:
            if not _spec_ok(spec):
                raise ConvertError(
                    f"转换 {src_ext} → {dst_ext} 缺少依赖：{', '.join(spec.requires)}，请先 pip install"
                )
            dst = Path(dst_path) if dst_path else unique_path(src.with_suffix("." + dst_ext))
            dst.parent.mkdir(parents=True, exist_ok=True)
            try:
                spec.func(src, dst, **opts)
            except ConvertError:
                raise
            except Exception as e:
                raise ConvertError(f"转换失败 {src.name} → {dst_ext}: {e}") from e
            return dst
    raise ConvertError(f"不支持的转换：.{src_ext} → .{dst_ext}")
```

同时创建空占位 `file_converter/converters/{documents,images,pdf,textdata}.py`（内容先为 `# 见后续任务`，本任务 Step 5 之后再由后续任务填充——为使骨架测试通过，`converters/__init__.py` 临时只导入 `base`：`from . import base  # noqa: F401`，Task 2-5 每完成一个模块就把它加回导入列表）。

- [ ] **Step 5: 运行测试确认通过**

Run: `python -m pytest tests/test_registry.py -v`
Expected: 5 passed

- [ ] **Step 6: Commit**

```bash
git add -A && git commit -m "feat: 项目骨架 — 注册表/依赖探测/基础类型"
```

---

### Task 2: 文本与数据转换（textdata）

**Files:**
- Create: `file_converter/converters/textdata.py`
- Modify: `file_converter/converters/__init__.py`（导入加 `textdata`）
- Test: `tests/test_textdata.py`

**Interfaces:**
- Consumes: `register`, `ConvertError`（Task 1）
- Produces: 注册 csv→json/yaml/xlsx 之外的 json↔yaml、json→csv、yaml→csv、xml→json、md→html、html→txt、txt→txt(编码)。`txt` 编码转换签名：`(src, dst, src_encoding=None, dst_encoding="utf-8", **opts)`。

- [ ] **Step 1: 写失败测试 `tests/test_textdata.py`**

```python
import json

import pytest
import yaml

from file_converter import registry


def test_csv_to_json(tmp_path):
    src = tmp_path / "a.csv"
    src.write_text("name,age\n张三,30\n李四,25\n", encoding="utf-8-sig")
    out = registry.convert(src, "json")
    data = json.loads(out.read_text(encoding="utf-8"))
    assert data == [{"name": "张三", "age": "30"}, {"name": "李四", "age": "25"}]


def test_json_to_csv(tmp_path):
    src = tmp_path / "a.json"
    src.write_text(json.dumps([{"a": 1, "b": 2}, {"a": 3, "b": 4}], ensure_ascii=False), encoding="utf-8")
    out = registry.convert(src, "csv")
    text = out.read_text(encoding="utf-8-sig")
    assert "a,b" in text and "1,2" in text and "3,4" in text


def test_json_to_csv_requires_records(tmp_path):
    src = tmp_path / "bad.json"
    src.write_text('{"a": 1}', encoding="utf-8")
    from file_converter.converters.base import ConvertError

    with pytest.raises(ConvertError, match="记录"):
        registry.convert(src, "csv")


def test_json_yaml_roundtrip(tmp_path):
    src = tmp_path / "a.json"
    src.write_text('{"x": [1, 2], "y": "中文"}', encoding="utf-8")
    yml = registry.convert(src, "yaml")
    assert yaml.safe_load(yml.read_text(encoding="utf-8")) == {"x": [1, 2], "y": "中文"}
    back = registry.convert(yml, "json")
    assert json.loads(back.read_text(encoding="utf-8"))["y"] == "中文"


def test_yaml_to_csv(tmp_path):
    src = tmp_path / "a.yaml"
    src.write_text("- a: 1\n  b: 2\n- a: 3\n  b: 4\n", encoding="utf-8")
    out = registry.convert(src, "csv")
    assert "a,b" in out.read_text(encoding="utf-8-sig")


def test_xml_to_json(tmp_path):
    src = tmp_path / "a.xml"
    src.write_text("<root><item id='1'>hi</item><item id='2'>yo</item></root>", encoding="utf-8")
    out = registry.convert(src, "json")
    data = json.loads(out.read_text(encoding="utf-8"))
    assert data["root"]["item"][0]["@id"] == "1"
    assert data["root"]["item"][0]["#text"] == "hi"


def test_md_to_html(tmp_path):
    src = tmp_path / "a.md"
    src.write_text("# 标题\n\n正文 **加粗**\n", encoding="utf-8")
    out = registry.convert(src, "html")
    html = out.read_text(encoding="utf-8")
    assert "<h1>" in html and "<strong>" in html


def test_html_to_txt(tmp_path):
    src = tmp_path / "a.html"
    src.write_text("<html><body><h1>标题</h1><p>段落</p></body></html>", encoding="utf-8")
    out = registry.convert(src, "txt")
    text = out.read_text(encoding="utf-8")
    assert "标题" in text and "段落" in text and "<h1>" not in text


def test_txt_encoding_gbk_to_utf8(tmp_path):
    src = tmp_path / "a.txt"
    src.write_bytes("中文编码测试".encode("gbk"))
    out = registry.convert(src, "txt", dst_encoding="utf-8")
    assert out.read_text(encoding="utf-8") == "中文编码测试"
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_textdata.py -v`
Expected: FAIL（不支持 / 模块不存在）

- [ ] **Step 3: 实现 `file_converter/converters/textdata.py`**

```python
import csv
import json
import xml.etree.ElementTree as ET
from pathlib import Path

from .base import ConvertError, register

_ENCODINGS_TRY = ("utf-8", "gbk", "big5", "shift_jis", "latin-1")


def _read_text_auto(path: Path, encoding: str | None = None) -> str:
    raw = path.read_bytes()
    if encoding:
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError as e:
            raise ConvertError(f"无法用 {encoding} 解码 {path.name}: {e}")
    for enc in _ENCODINGS_TRY:
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise ConvertError(f"无法识别 {path.name} 的文本编码")


def _records_from_csv(path: Path, encoding: str | None) -> list[dict]:
    text = _read_text_auto(path, encoding)
    rows = list(csv.reader(text.splitlines()))
    if not rows:
        raise ConvertError(f"{path.name} 是空 CSV")
    header = rows[0]
    return [dict(zip(header, r)) for r in rows[1:] if any(c.strip() for c in r)]


def _records_to_csv(records: list[dict], dst: Path) -> None:
    fields: list[str] = []
    for rec in records:
        for k in rec:
            if k not in fields:
                fields.append(k)
    with open(dst, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(records)


def _load_records(obj, name: str) -> list[dict]:
    if isinstance(obj, list) and all(isinstance(x, dict) for x in obj):
        return obj
    raise ConvertError(f"{name} 转 CSV 要求内容是由对象组成的记录数组（如 [{{\"a\":1}}]）")


@register("csv", ["json"])
def csv_to_json(src: Path, dst: Path, src_encoding=None, **opts):
    records = _records_from_csv(src, src_encoding)
    dst.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")


@register("csv", ["yaml", "yml"], requires=("yaml",))
def csv_to_yaml(src: Path, dst: Path, src_encoding=None, **opts):
    import yaml

    records = _records_from_csv(src, src_encoding)
    dst.write_text(yaml.safe_dump(records, allow_unicode=True, sort_keys=False), encoding="utf-8")


@register("json", ["csv"])
def json_to_csv(src: Path, dst: Path, **opts):
    data = json.loads(_read_text_auto(src))
    _records_to_csv(_load_records(data, src.name), dst)


@register("json", ["yaml", "yml"], requires=("yaml",))
def json_to_yaml(src: Path, dst: Path, **opts):
    import yaml

    data = json.loads(_read_text_auto(src))
    dst.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding="utf-8")


@register(["yaml", "yml"], ["json"], requires=("yaml",))
def yaml_to_json(src: Path, dst: Path, **opts):
    import yaml

    data = yaml.safe_load(_read_text_auto(src))
    dst.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


@register(["yaml", "yml"], ["csv"], requires=("yaml",))
def yaml_to_csv(src: Path, dst: Path, **opts):
    import yaml

    data = yaml.safe_load(_read_text_auto(src))
    _records_to_csv(_load_records(data, src.name), dst)


def _elem_to_dict(elem: ET.Element):
    node: dict = {}
    for k, v in elem.attrib.items():
        node[f"@{k}"] = v
    children = list(elem)
    if children:
        grouped: dict[str, list] = {}
        for child in children:
            grouped.setdefault(child.tag, []).append(_elem_to_dict(child))
        for tag, items in grouped.items():
            node[tag] = items[0] if len(items) == 1 else items
    text = (elem.text or "").strip()
    if text:
        node["#text"] = text
    return node


@register("xml", ["json"])
def xml_to_json(src: Path, dst: Path, **opts):
    try:
        root = ET.fromstring(_read_text_auto(src))
    except ET.ParseError as e:
        raise ConvertError(f"XML 解析失败 {src.name}: {e}")
    dst.write_text(
        json.dumps({root.tag: _elem_to_dict(root)}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


@register(["md", "markdown"], ["html"], requires=("markdown",))
def md_to_html(src: Path, dst: Path, **opts):
    import markdown

    body = markdown.markdown(_read_text_auto(src), extensions=["tables", "fenced_code"])
    dst.write_text(
        '<!DOCTYPE html>\n<html><head><meta charset="utf-8"></head><body>\n'
        + body
        + "\n</body></html>\n",
        encoding="utf-8",
    )


@register(["html", "htm"], ["txt"], requires=("bs4",))
def html_to_txt(src: Path, dst: Path, **opts):
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(_read_text_auto(src), "html.parser")
    dst.write_text(soup.get_text("\n", strip=True), encoding="utf-8")


@register("txt", ["txt"], label="文本编码转换")
def txt_reencode(src: Path, dst: Path, src_encoding=None, dst_encoding="utf-8", **opts):
    text = _read_text_auto(src, src_encoding)
    try:
        dst.write_bytes(text.encode(dst_encoding))
    except (LookupError, UnicodeEncodeError) as e:
        raise ConvertError(f"无法编码为 {dst_encoding}: {e}")
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_textdata.py -v`
Expected: 9 passed

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: 文本与数据转换 (csv/json/yaml/xml/md/html/编码)"
```

---

### Task 3: 图片转换（images）

**Files:**
- Create: `file_converter/converters/images.py`
- Modify: `file_converter/converters/__init__.py`（导入加 `images`）
- Test: `tests/test_images.py`

**Interfaces:**
- Consumes: `register`, `ConvertError`
- Produces: 注册 png/jpg/jpeg/webp/bmp/gif/tiff/ico 互转 与 这些格式→pdf。

- [ ] **Step 1: 写失败测试 `tests/test_images.py`**

```python
from pathlib import Path

import pytest
from PIL import Image

from file_converter import registry


@pytest.fixture()
def png_file(tmp_path):
    p = tmp_path / "pic.png"
    Image.new("RGB", (40, 30), (200, 30, 30)).save(p)
    return p


def test_png_to_jpg(png_file):
    out = registry.convert(png_file, "jpg")
    with Image.open(out) as im:
        assert im.format == "JPEG" and im.size == (40, 30)


def test_png_to_webp(png_file):
    out = registry.convert(png_file, "webp")
    with Image.open(out) as im:
        assert im.format == "WEBP"


def test_rgba_to_jpg_flattens(tmp_path):
    p = tmp_path / "t.png"
    Image.new("RGBA", (10, 10), (0, 0, 255, 128)).save(p)
    out = registry.convert(p, "jpg")
    with Image.open(out) as im:
        assert im.mode == "RGB"


def test_image_to_pdf(png_file):
    out = registry.convert(png_file, "pdf")
    assert out.read_bytes().startswith(b"%PDF")


def test_targets_include_image_formats(png_file):
    targets = registry.available_targets("png")
    for t in ("jpg", "webp", "bmp", "gif", "tiff", "ico", "pdf"):
        assert t in targets
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_images.py -v`
Expected: FAIL

- [ ] **Step 3: 实现 `file_converter/converters/images.py`**

```python
from pathlib import Path

from .base import ConvertError, register

IMG_FORMATS = ["png", "jpg", "jpeg", "webp", "bmp", "gif", "tiff", "ico"]


@register(IMG_FORMATS, IMG_FORMATS, requires=("PIL",), label="图片格式互转")
def image_convert(src: Path, dst: Path, **opts):
    from PIL import Image

    dst_ext = dst.suffix.lower()
    with Image.open(src) as im:
        if dst_ext in (".jpg", ".jpeg") and im.mode not in ("RGB", "L", "CMYK"):
            im = im.convert("RGB")
        elif dst_ext == ".png" and im.mode == "P":
            im = im.convert("RGBA")
        elif dst_ext == ".ico":
            im = im.convert("RGBA")
            im.thumbnail((256, 256))
        im.save(dst)


@register(IMG_FORMATS, ["pdf"], requires=("PIL",), label="图片→PDF")
def image_to_pdf(src: Path, dst: Path, **opts):
    from PIL import Image

    with Image.open(src) as im:
        if im.mode != "RGB":
            im = im.convert("RGB")
        im.save(dst, "PDF", resolution=100.0)
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_images.py -v`
Expected: 5 passed

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: 图片互转与图片转PDF"
```

---

### Task 4: Office 文档转换（documents）

**Files:**
- Create: `file_converter/converters/documents.py`
- Modify: `file_converter/converters/__init__.py`（导入加 `documents`）
- Test: `tests/test_documents.py`

**Interfaces:**
- Consumes: `register`, `ConvertError`, `deps.office_available()`, `deps.libreoffice_path()`
- Produces: 注册 docx→txt/md/pdf，xlsx→csv/pdf，csv→xlsx，pptx→txt/pdf。注意 csv→xlsx 与 textdata 的 csv→json/yaml 并存不冲突（不同 dst）。

- [ ] **Step 1: 写失败测试 `tests/test_documents.py`**

```python
import csv

import pytest

from file_converter import registry


def test_docx_to_txt_and_md(tmp_path):
    import docx

    d = docx.Document()
    d.add_heading("报告标题", level=1)
    d.add_paragraph("第一段内容")
    p = tmp_path / "a.docx"
    d.save(str(p))

    txt = registry.convert(p, "txt")
    assert "第一段内容" in txt.read_text(encoding="utf-8")

    md = registry.convert(p, "md")
    md_text = md.read_text(encoding="utf-8")
    assert md_text.startswith("# 报告标题")


def test_xlsx_to_csv(tmp_path):
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "数据"
    ws.append(["姓名", "分数"])
    ws.append(["张三", 90])
    p = tmp_path / "a.xlsx"
    wb.save(str(p))

    out = registry.convert(p, "csv")
    rows = list(csv.reader(out.read_text(encoding="utf-8-sig").splitlines()))
    assert rows == [["姓名", "分数"], ["张三", "90"]]


def test_xlsx_to_csv_multisheet(tmp_path):
    import openpyxl

    wb = openpyxl.Workbook()
    wb.active.title = "一"
    wb.create_sheet("二")
    wb["一"].append(["a"])
    wb["二"].append(["b"])
    p = tmp_path / "m.xlsx"
    wb.save(str(p))

    out = registry.convert(p, "csv")
    assert out.name == "m_一.csv"
    assert (tmp_path / "m_二.csv").exists()


def test_csv_to_xlsx(tmp_path):
    src = tmp_path / "a.csv"
    src.write_text("x,y\n1,2\n", encoding="utf-8-sig")
    out = registry.convert(src, "xlsx")

    import openpyxl

    wb = openpyxl.load_workbook(str(out))
    ws = wb.active
    assert ws["A1"].value == "x" and ws["B2"].value == "2"


def test_pptx_to_txt(tmp_path):
    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "演示标题"
    slide.placeholders[1].text = "要点一"
    p = tmp_path / "a.pptx"
    prs.save(str(p))

    out = registry.convert(p, "txt")
    text = out.read_text(encoding="utf-8")
    assert "演示标题" in text and "要点一" in text
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_documents.py -v`
Expected: FAIL

- [ ] **Step 3: 实现 `file_converter/converters/documents.py`**

```python
import csv
import subprocess
from pathlib import Path

from .. import deps
from .base import ConvertError, register
from .textdata import _read_text_auto


@register("docx", ["txt"], requires=("docx",), label="Word→纯文本")
def docx_to_txt(src: Path, dst: Path, **opts):
    import docx

    document = docx.Document(str(src))
    lines = [p.text for p in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            lines.append("\t".join(cell.text for cell in row.cells))
    dst.write_text("\n".join(lines), encoding="utf-8")


@register("docx", ["md"], requires=("docx",), label="Word→Markdown")
def docx_to_md(src: Path, dst: Path, **opts):
    import docx

    document = docx.Document(str(src))
    lines = []
    for p in document.paragraphs:
        style = (p.style.name or "") if p.style else ""
        text = p.text
        if style.startswith("Heading"):
            try:
                level = int(style.split()[-1])
            except ValueError:
                level = 1
            lines.append("#" * min(max(level, 1), 6) + " " + text)
        else:
            lines.append(text)
    dst.write_text("\n\n".join(lines) + "\n", encoding="utf-8")


@register("xlsx", ["csv"], requires=("openpyxl",), label="Excel→CSV")
def xlsx_to_csv(src: Path, dst: Path, **opts):
    import openpyxl

    wb = openpyxl.load_workbook(str(src), data_only=True)
    names = wb.sheetnames
    for name in names:
        ws = wb[name]
        target = dst if len(names) == 1 else dst.with_name(f"{dst.stem}_{name}{dst.suffix}")
        with open(target, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            for row in ws.iter_rows(values_only=True):
                w.writerow(["" if v is None else v for v in row])


@register("csv", ["xlsx"], requires=("openpyxl",), label="CSV→Excel")
def csv_to_xlsx(src: Path, dst: Path, src_encoding=None, **opts):
    import openpyxl

    text = _read_text_auto(src, src_encoding)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = src.stem[:31] or "Sheet1"
    for row in csv.reader(text.splitlines()):
        ws.append(row)
    wb.save(str(dst))


@register("pptx", ["txt"], requires=("pptx",), label="PPT→纯文本")
def pptx_to_txt(src: Path, dst: Path, **opts):
    from pptx import Presentation

    prs = Presentation(str(src))
    lines: list[str] = []
    for i, slide in enumerate(prs.slides, 1):
        lines.append(f"=== 第 {i} 页 ===")
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    t = "".join(run.text for run in para.runs)
                    if t.strip():
                        lines.append(t)
    dst.write_text("\n".join(lines) + "\n", encoding="utf-8")


_COM_APPS = {
    ".docx": "Word.Application",
    ".xlsx": "Excel.Application",
    ".pptx": "PowerPoint.Application",
}


def _com_to_pdf(src: Path, dst: Path) -> None:
    import pythoncom
    import win32com.client

    ext = src.suffix.lower()
    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx(_COM_APPS[ext])
        if ext != ".pptx":  # PowerPoint 不允许 Visible=False
            app.Visible = False
        if ext == ".docx":
            doc = app.Documents.Open(str(src.resolve()), ReadOnly=True)
            doc.SaveAs2(str(dst.resolve()), FileFormat=17)  # wdFormatPDF
            doc.Close(False)
        elif ext == ".xlsx":
            wb = app.Workbooks.Open(str(src.resolve()), ReadOnly=True)
            wb.ExportAsFixedFormat(0, str(dst.resolve()))  # xlTypePDF = 0
            wb.Close(False)
        else:
            pres = app.Presentations.Open(str(src.resolve()), WithWindow=False)
            pres.SaveAs(str(dst.resolve()), 32)  # ppSaveAsPDF
            pres.Close()
    except Exception as e:
        raise ConvertError(f"Office 转换 PDF 失败 {src.name}: {e}")
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _libreoffice_to_pdf(src: Path, dst: Path) -> None:
    exe = deps.libreoffice_path()
    dst.parent.mkdir(parents=True, exist_ok=True)
    try:
        r = subprocess.run(
            [exe, "--headless", "--convert-to", "pdf", "--outdir", str(dst.parent), str(src)],
            capture_output=True, text=True, timeout=300,
        )
    except subprocess.TimeoutExpired:
        raise ConvertError(f"LibreOffice 转换超时 {src.name}")
    produced = dst.parent / (src.stem + ".pdf")
    if r.returncode != 0 or not produced.exists():
        raise ConvertError(f"LibreOffice 转换失败 {src.name}: {(r.stderr or r.stdout).strip()}")
    if produced != dst:
        produced.replace(dst)


@register(["docx", "xlsx", "pptx"], ["pdf"], label="Office→PDF")
def office_to_pdf(src: Path, dst: Path, **opts):
    if deps.office_available():
        _com_to_pdf(src, dst)
    elif deps.libreoffice_path():
        _libreoffice_to_pdf(src, dst)
    else:
        raise ConvertError("转换为 PDF 需要安装 MS Office 或 LibreOffice")
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_documents.py -v`
Expected: 5 passed

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: Office 文档转换 (docx/xlsx/pptx/csv + COM/LibreOffice PDF)"
```

---

### Task 5: PDF 工具（pdf）

**Files:**
- Create: `file_converter/converters/pdf.py`
- Modify: `file_converter/converters/__init__.py`（导入加 `pdf`，恢复为 `from . import base, documents, images, pdf, textdata`）
- Test: `tests/test_pdf.py`

**Interfaces:**
- Consumes: `register`, `ConvertError`
- Produces: 注册 pdf→txt/png/jpg；模块级函数供 CLI/GUI 直接调用：
  - `merge_pdfs(srcs: list[Path], dst: Path) -> None`
  - `split_pdf(src: Path, dst_dir: Path, pages: str) -> list[Path]`（`pages` 形如 `"1-3,5"`，按连续段拆成多个 PDF）
  - `rotate_pdf(src: Path, dst: Path, angle: int = 90) -> None`
  - `parse_page_spec(spec: str, page_count: int) -> list[list[int]]`（返回连续段，每段为升序页码列表，1-based）

- [ ] **Step 1: 写失败测试 `tests/test_pdf.py`**

```python
import fitz
import pytest
from pypdf import PdfReader

from file_converter import registry
from file_converter.converters.base import ConvertError
from file_converter.converters.pdf import (
    merge_pdfs,
    parse_page_spec,
    rotate_pdf,
    split_pdf,
)


def _make_pdf(path, texts):
    doc = fitz.open()
    for t in texts:
        page = doc.new_page()
        page.insert_text((72, 72), t)
    doc.save(str(path))
    doc.close()


@pytest.fixture()
def pdf3(tmp_path):
    p = tmp_path / "a.pdf"
    _make_pdf(p, ["第一页", "第二页", "第三页"])
    return p


def test_parse_page_spec():
    assert parse_page_spec("1-3,5", 10) == [[1, 2, 3], [5]]
    assert parse_page_spec("2,1", 5) == [[1, 2]]
    with pytest.raises(ConvertError):
        parse_page_spec("9", 3)
    with pytest.raises(ConvertError):
        parse_page_spec("a-b", 3)


def test_pdf_to_txt(pdf3):
    out = registry.convert(pdf3, "txt")
    text = out.read_text(encoding="utf-8")
    assert "第一页" in text and "第三页" in text


def test_pdf_to_png(pdf3):
    out = registry.convert(pdf3, "png")
    assert out.name == "a_p1.png"
    assert (pdf3.parent / "a_p3.png").exists()
    assert out.read_bytes()[:8] == b"\x89PNG\r\n\x1a\n"


def test_merge(pdf3, tmp_path):
    b = tmp_path / "b.pdf"
    _make_pdf(b, ["乙一"])
    out = tmp_path / "merged.pdf"
    merge_pdfs([pdf3, b], out)
    assert len(PdfReader(str(out)).pages) == 4


def test_split(pdf3, tmp_path):
    outs = split_pdf(pdf3, tmp_path / "out", "1-2,3")
    assert len(outs) == 2
    assert [len(PdfReader(str(o)).pages) for o in outs] == [2, 1]


def test_rotate(pdf3, tmp_path):
    out = tmp_path / "r.pdf"
    rotate_pdf(pdf3, out, 90)
    assert PdfReader(str(out)).pages[0].rotation == 90
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_pdf.py -v`
Expected: FAIL

- [ ] **Step 3: 实现 `file_converter/converters/pdf.py`**

```python
from pathlib import Path

from .base import ConvertError, register


@register("pdf", ["txt"], requires=("pypdf",), label="PDF→纯文本")
def pdf_to_txt(src: Path, dst: Path, **opts):
    from pypdf import PdfReader

    reader = PdfReader(str(src))
    parts = [(page.extract_text() or "") for page in reader.pages]
    dst.write_text("\n\n".join(parts), encoding="utf-8")


def _pdf_to_images(src: Path, dst: Path, dpi: int = 150):
    import fitz

    doc = fitz.open(str(src))
    try:
        zoom = dpi / 72
        matrix = fitz.Matrix(zoom, zoom)
        single = len(doc) == 1
        for i, page in enumerate(doc):
            target = dst if single else dst.with_name(f"{dst.stem}_p{i + 1}{dst.suffix}")
            page.get_pixmap(matrix=matrix).save(str(target))
    finally:
        doc.close()


@register("pdf", ["png"], requires=("fitz",), label="PDF→PNG（逐页）")
def pdf_to_png(src: Path, dst: Path, dpi: int = 150, **opts):
    _pdf_to_images(src, dst, dpi)


@register("pdf", ["jpg", "jpeg"], requires=("fitz",), label="PDF→JPG（逐页）")
def pdf_to_jpg(src: Path, dst: Path, dpi: int = 150, **opts):
    _pdf_to_images(src, dst, dpi)


def parse_page_spec(spec: str, page_count: int) -> list[list[int]]:
    """把 '1-3,5' 解析成连续段 [[1,2,3],[5]]（1-based）。非法输入抛 ConvertError。"""
    pages: set[int] = set()
    try:
        for part in spec.split(","):
            part = part.strip()
            if "-" in part:
                a, b = part.split("-", 1)
                lo, hi = int(a), int(b)
                if lo > hi:
                    raise ValueError
                pages.update(range(lo, hi + 1))
            else:
                pages.add(int(part))
    except ValueError:
        raise ConvertError(f"页码格式无效：{spec!r}（示例：1-3,5）")
    if not pages or min(pages) < 1 or max(pages) > page_count:
        raise ConvertError(f"页码超出范围（共 {page_count} 页）：{spec!r}")
    ordered = sorted(pages)
    segments: list[list[int]] = [[ordered[0]]]
    for p in ordered[1:]:
        if p == segments[-1][-1] + 1:
            segments[-1].append(p)
        else:
            segments.append([p])
    return segments


def merge_pdfs(srcs: list[Path], dst: Path) -> None:
    from pypdf import PdfWriter

    if len(srcs) < 2:
        raise ConvertError("合并 PDF 至少需要 2 个文件")
    writer = PdfWriter()
    try:
        for s in srcs:
            writer.append(str(s))
        dst.parent.mkdir(parents=True, exist_ok=True)
        with open(dst, "wb") as f:
            writer.write(f)
    except ConvertError:
        raise
    except Exception as e:
        raise ConvertError(f"合并 PDF 失败: {e}")
    finally:
        writer.close()


def split_pdf(src: Path, dst_dir: Path, pages: str) -> list[Path]:
    from pypdf import PdfReader, PdfWriter

    from ..registry import unique_path

    reader = PdfReader(str(src))
    segments = parse_page_spec(pages, len(reader.pages))
    dst_dir.mkdir(parents=True, exist_ok=True)
    outs: list[Path] = []
    for seg in segments:
        writer = PdfWriter()
        for p in seg:
            writer.add_page(reader.pages[p - 1])
        name = f"{src.stem}_p{seg[0]}-{seg[-1]}.pdf" if len(seg) > 1 else f"{src.stem}_p{seg[0]}.pdf"
        out = unique_path(dst_dir / name)
        with open(out, "wb") as f:
            writer.write(f)
        writer.close()
        outs.append(out)
    return outs


def rotate_pdf(src: Path, dst: Path, angle: int = 90) -> None:
    from pypdf import PdfReader, PdfWriter

    if angle not in (90, 180, 270):
        raise ConvertError(f"旋转角度必须是 90/180/270，收到 {angle}")
    reader = PdfReader(str(src))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page.rotate(angle))
    dst.parent.mkdir(parents=True, exist_ok=True)
    with open(dst, "wb") as f:
        writer.write(f)
    writer.close()
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_pdf.py tests/test_registry.py -v`
Expected: 全部 passed（registry 测试确认无循环导入）

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: PDF 转换与工具 (txt/png/jpg/合并/拆分/旋转)"
```

---

### Task 6: CLI（__main__.py）

**Files:**
- Create: `file_converter/__main__.py`
- Test: `tests/test_cli.py`

**Interfaces:**
- Consumes: `registry.convert/available_targets/matrix`、`converters.pdf.merge_pdfs/split_pdf/rotate_pdf`、`registry.unique_path`、`deps`
- Produces: `main(argv=None) -> int`；子命令 `convert`（可省略）、`merge`、`split`、`rotate`、`list`；`--gui` 启动图形界面。

- [ ] **Step 1: 写失败测试 `tests/test_cli.py`**

```python
import json

from file_converter.__main__ import main


def test_list_runs(capsys):
    assert main(["list"]) == 0
    out = capsys.readouterr().out
    assert "docx" in out and "png" in out


def test_convert_implicit_subcommand(tmp_path, capsys):
    src = tmp_path / "a.csv"
    src.write_text("k,v\n1,2\n", encoding="utf-8")
    assert main([str(src), "-t", "json", "-o", str(tmp_path)]) == 0
    assert json.loads((tmp_path / "a.json").read_text(encoding="utf-8")) == [{"k": "1", "v": "2"}]


def test_convert_failure_continues(tmp_path, capsys):
    good = tmp_path / "g.csv"
    good.write_text("a\n1\n", encoding="utf-8")
    bad = tmp_path / "b.csv"
    bad.write_text("a\n1\n", encoding="utf-8")
    bad.rename(tmp_path / "b.nope")  # 不支持的扩展名
    rc = main([str(good), str(tmp_path / "b.nope"), "-t", "json", "-o", str(tmp_path)])
    assert rc == 1  # 有失败
    assert (tmp_path / "g.json").exists()  # 但成功的照常产出
    assert "失败" in capsys.readouterr().out


def test_merge_split_rotate(tmp_path):
    import fitz
    from pypdf import PdfReader

    def mk(name, n):
        doc = fitz.open()
        for i in range(n):
            page = doc.new_page()
            page.insert_text((72, 72), f"{name}{i}")
        doc.save(str(tmp_path / name))
        doc.close()

    mk("m1.pdf", 2)
    mk("m2.pdf", 1)
    merged = tmp_path / "merged.pdf"
    assert main(["merge", str(tmp_path / "m1.pdf"), str(tmp_path / "m2.pdf"), "-o", str(merged)]) == 0
    assert len(PdfReader(str(merged)).pages) == 3

    outdir = tmp_path / "parts"
    assert main(["split", str(merged), "--pages", "1-2,3", "-o", str(outdir)]) == 0
    assert len(list(outdir.glob("*.pdf"))) == 2

    rot = tmp_path / "rot.pdf"
    assert main(["rotate", str(tmp_path / "m2.pdf"), "--angle", "180", "-o", str(rot)]) == 0
    assert PdfReader(str(rot)).pages[0].rotation == 180
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_cli.py -v`
Expected: FAIL

- [ ] **Step 3: 实现 `file_converter/__main__.py`**

```python
import argparse
import sys
from pathlib import Path

from . import deps, registry
from .converters.base import ConvertError
from .converters.pdf import merge_pdfs, rotate_pdf, split_pdf

_SUBCOMMANDS = ("convert", "merge", "split", "rotate", "list")


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="file_converter", description="万能文件转换工具（CLI）")
    p.add_argument("--gui", action="store_true", help="启动图形界面")
    sub = p.add_subparsers(dest="cmd")

    c = sub.add_parser("convert", help="格式转换（默认子命令，可省略）")
    c.add_argument("files", nargs="+", help="输入文件（可多个）")
    c.add_argument("-t", "--to", required=True, help="目标格式，如 pdf / jpg / xlsx")
    c.add_argument("-o", "--outdir", help="输出目录（默认与源文件同目录）")
    c.add_argument("--src-encoding", help="源文本编码（默认自动探测）")
    c.add_argument("--dst-encoding", default="utf-8", help="txt 输出编码（默认 utf-8）")
    c.add_argument("--dpi", type=int, default=150, help="PDF 转图片的 DPI（默认 150）")

    m = sub.add_parser("merge", help="合并多个 PDF")
    m.add_argument("files", nargs="+", help="至少 2 个 PDF")
    m.add_argument("-o", "--output", required=True, help="输出 PDF 路径")

    s = sub.add_parser("split", help="按页码拆分 PDF")
    s.add_argument("file", help="输入 PDF")
    s.add_argument("--pages", required=True, help="页码范围，如 1-3,5")
    s.add_argument("-o", "--outdir", required=True, help="输出目录")

    r = sub.add_parser("rotate", help="旋转 PDF 页面")
    r.add_argument("file", help="输入 PDF")
    r.add_argument("--angle", type=int, default=90, choices=(90, 180, 270), help="旋转角度")
    r.add_argument("-o", "--output", required=True, help="输出 PDF 路径")

    sub.add_parser("list", help="列出全部支持的转换")
    return p


def _cmd_list() -> int:
    print("支持的转换（[*] 表示依赖缺失，暂不可用）：\n")
    for srcs, dsts, label, ok in registry.matrix():
        mark = "" if ok else " [*]"
        print(f"  {', '.join(srcs):30s} → {', '.join(dsts):20s} {label}{mark}")
    if deps.office_available():
        print("\nOffice→PDF 引擎：MS Office (COM)")
    elif deps.libreoffice_path():
        print("\nOffice→PDF 引擎：LibreOffice")
    else:
        print("\nOffice→PDF 引擎：不可用（需安装 MS Office 或 LibreOffice）")
    return 0


def _cmd_convert(args) -> int:
    ok_count, fail_count = 0, 0
    for f in args.files:
        src = Path(f)
        if not src.exists():
            print(f"[失败] {f}: 文件不存在")
            fail_count += 1
            continue
        dst = None
        if args.outdir:
            dst = registry.unique_path(Path(args.outdir) / (src.stem + "." + args.to.lower().lstrip(".")))
        try:
            out = registry.convert(
                src, args.to, dst,
                src_encoding=args.src_encoding,
                dst_encoding=args.dst_encoding,
                dpi=args.dpi,
            )
            print(f"[成功] {src.name} → {out}")
            ok_count += 1
        except ConvertError as e:
            print(f"[失败] {src.name}: {e}")
            fail_count += 1
    print(f"\n完成：成功 {ok_count} 个，失败 {fail_count} 个")
    return 0 if fail_count == 0 else 1


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    parser = _build_parser()
    if argv and not argv[0].startswith("-") and argv[0] not in _SUBCOMMANDS:
        argv.insert(0, "convert")  # 省略子命令时按 convert 处理
    args = parser.parse_args(argv)

    if getattr(args, "gui", False) or args.cmd is None:
        if args.cmd is None and not getattr(args, "gui", False):
            parser.print_help()
            return 0
        from .gui import run

        run()
        return 0
    if args.cmd == "list":
        return _cmd_list()
    if args.cmd == "convert":
        return _cmd_convert(args)
    if args.cmd == "merge":
        try:
            merge_pdfs([Path(f) for f in args.files], Path(args.output))
            print(f"[成功] 已合并 → {args.output}")
            return 0
        except ConvertError as e:
            print(f"[失败] {e}")
            return 1
    if args.cmd == "split":
        try:
            outs = split_pdf(Path(args.file), Path(args.outdir), args.pages)
            for o in outs:
                print(f"[成功] {o}")
            return 0
        except ConvertError as e:
            print(f"[失败] {e}")
            return 1
    if args.cmd == "rotate":
        try:
            rotate_pdf(Path(args.file), Path(args.output), args.angle)
            print(f"[成功] 已旋转 {args.angle}° → {args.output}")
            return 0
        except ConvertError as e:
            print(f"[失败] {e}")
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/ -v`
Expected: 全部 passed

- [ ] **Step 5: 真实 CLI 冒烟**

```bash
python -m file_converter list
echo "name,score" > /tmp/smoke.csv && echo "张三,90" >> /tmp/smoke.csv
python -m file_converter /tmp/smoke.csv -t xlsx -o /tmp
python -m file_converter /tmp/smoke.csv -t json -o /tmp
```
Expected: 三条命令均成功，`/tmp/smoke.xlsx`、`/tmp/smoke.json` 生成

- [ ] **Step 6: Commit**

```bash
git add -A && git commit -m "feat: CLI (convert/merge/split/rotate/list/--gui)"
```

---

### Task 7: Tkinter GUI（gui.py）

**Files:**
- Create: `file_converter/gui.py`
- Test: 无（GUI 不测；冒烟 = 能 import、能实例化后立即 destroy）

**Interfaces:**
- Consumes: `registry.available_targets/convert/unique_path`、`converters.pdf.merge_pdfs/split_pdf/rotate_pdf`、`deps`
- Produces: `run() -> None`（启动主循环）；`ConverterApp(tk.Tk)`。

- [ ] **Step 1: 实现 `file_converter/gui.py`**

要点：文件列表 Listbox；添加文件/文件夹、移除、清空；目标格式 Combobox（取列表中所有文件 `available_targets` 的交集，单文件时就是它的目标集）；输出目录 Entry+浏览（默认"与源文件同目录"）；选项区：DPI Spinbox、编码 Combobox（utf-8/gbk/big5/shift_jis）；开始转换按钮 → `threading.Thread` 执行，主线程 `after(100, poll)` 从 `queue.Queue` 取进度更新 ttk.Progressbar 与日志 Text；结束后 `messagebox.showinfo` 汇总。第二行按钮：合并 PDF（用当前列表）、拆分 PDF（`simpledialog.askstring` 问页码）、旋转 PDF（问角度）。窗口标题"万能文件转换工具"，最小尺寸 760×520。

```python
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
```

- [ ] **Step 2: 冒烟验证（能实例化）**

```bash
python -c "import tkinter as tk; from file_converter.gui import ConverterApp; r = tk.Tk(); ConverterApp(r); r.update(); r.destroy(); print('GUI OK')"
```
Expected: 输出 `GUI OK`

- [ ] **Step 3: Commit**

```bash
git add -A && git commit -m "feat: Tkinter 图形界面"
```

---

### Task 8: README + 端到端验收

**Files:**
- Create: `README.md`
- Modify: 无

- [ ] **Step 1: 写 `README.md`**

包含：功能简介、安装（`pip install -r requirements.txt`）、GUI 启动（`python -m file_converter --gui`）、CLI 用法示例（convert/merge/split/rotate/list）、支持的转换矩阵表、Office→PDF 引擎说明（MS Office 优先，LibreOffice 回退）、常见问题（GBK 编码、CSV 用 utf-8-sig）。

- [ ] **Step 2: 全量测试 + 端到端真实转换**

```bash
python -m pytest tests/ -v
python -m file_converter list
# 用 python 现场造 docx/xlsx/pptx/png/pdf/md 各一个到 %TEMP%\fc_e2e，逐一真实转换：
python - <<'EOF'
import os, tempfile
from pathlib import Path
tmp = Path(tempfile.gettempdir()) / "fc_e2e"
tmp.mkdir(exist_ok=True)
import docx, openpyxl
from pptx import Presentation
from PIL import Image
import fitz

d = docx.Document(); d.add_paragraph("端到端"); d.save(str(tmp / "t.docx"))
wb = openpyxl.Workbook(); wb.active.append(["a", "1"]); wb.save(str(tmp / "t.xlsx"))
prs = Presentation(); prs.slides.add_slide(prs.slide_layouts[0]); prs.save(str(tmp / "t.pptx"))
Image.new("RGB", (20, 20), (1, 2, 3)).save(tmp / "t.png")
doc = fitz.open(); pg = doc.new_page(); pg.insert_text((72, 72), "pdf 文本"); doc.save(str(tmp / "t.pdf")); doc.close()
(tmp / "t.md").write_text("# 标题", encoding="utf-8")
print(tmp)
EOF
python -m file_converter %TEMP%\fc_e2e\t.docx -t txt -o %TEMP%\fc_e2e\out
python -m file_converter %TEMP%\fc_e2e\t.xlsx -t csv -o %TEMP%\fc_e2e\out
python -m file_converter %TEMP%\fc_e2e\t.png -t jpg -o %TEMP%\fc_e2e\out
python -m file_converter %TEMP%\fc_e2e\t.png -t pdf -o %TEMP%\fc_e2e\out
python -m file_converter %TEMP%\fc_e2e\t.pdf -t png -o %TEMP%\fc_e2e\out
python -m file_converter %TEMP%\fc_e2e\t.md -t html -o %TEMP%\fc_e2e\out
# Office→PDF（本机有 MS Office）：
python -m file_converter %TEMP%\fc_e2e\t.docx -t pdf -o %TEMP%\fc_e2e\out
```
Expected: 全部 `[成功]`，`out` 目录产出对应文件；`pytest` 全绿。

- [ ] **Step 3: Commit**

```bash
git add -A && git commit -m "docs: README 与端到端验收"
```

---

## Self-Review 记录

- **Spec 覆盖**：转换矩阵每一行（docx/xlsx/pptx/csv/图片/PDF/数据/文本/编码）→ Task 2-5；注册表/依赖 → Task 1；CLI/GUI → Task 6/7；验收 → Task 8。PDF 合并/拆分/旋转在 Task 5 实现、Task 6/7 接出。
- **Placeholder 扫描**：无 TBD；每个代码步骤含完整代码。
- **类型一致性**：`register(src, dst, requires=(), label="")`、`registry.convert(src_path, dst_ext, dst_path=None, **opts) -> Path`、`unique_path(Path) -> Path`、`parse_page_spec(str, int) -> list[list[int]]` 在 Task 6/7 消费处签名一致；`textdata._read_text_auto(path, encoding)` 被 Task 4 `csv_to_xlsx` 复用。
