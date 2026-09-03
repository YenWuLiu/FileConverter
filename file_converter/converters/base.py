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
