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
