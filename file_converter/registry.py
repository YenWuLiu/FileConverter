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
