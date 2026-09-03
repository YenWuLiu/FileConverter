import csv
import json
import re
import xml.etree.ElementTree as ET
from pathlib import Path

from .base import ConvertError, register

_ENCODINGS_TRY = ("utf-8-sig", "gbk", "big5", "shift_jis", "latin-1")


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
    text = _read_text_auto(src)
    # 文本已解码，剥离 XML 声明里的 encoding 属性，否则 fromstring 拒绝带编码声明的 str
    text = re.sub(r"(<\?xml[^>]*?)\s*encoding\s*=\s*['\"][^'\"]*['\"]", r"\1", text, count=1)
    try:
        root = ET.fromstring(text)
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


@register("txt", ["txt"], label="文本编码转换", self_conversion=True)
def txt_reencode(src: Path, dst: Path, src_encoding=None, dst_encoding="utf-8", **opts):
    text = _read_text_auto(src, src_encoding)
    try:
        dst.write_bytes(text.encode(dst_encoding))
    except (LookupError, UnicodeEncodeError) as e:
        raise ConvertError(f"无法编码为 {dst_encoding}: {e}")
