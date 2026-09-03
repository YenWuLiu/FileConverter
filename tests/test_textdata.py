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


def test_xml_to_json_with_encoding_declaration(tmp_path):
    src = tmp_path / "decl.xml"
    src.write_text("<?xml version=\"1.0\" encoding=\"UTF-8\"?><root><a>你好</a></root>", encoding="utf-8")
    out = registry.convert(src, "json")
    data = json.loads(out.read_text(encoding="utf-8"))
    assert data["root"]["a"]["#text"] == "你好"


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
