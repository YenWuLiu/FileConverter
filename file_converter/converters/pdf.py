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
    segments: list[list[int]] = []
    for part in pages.split(","):
        segments.extend(parse_page_spec(part, len(reader.pages)))
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
