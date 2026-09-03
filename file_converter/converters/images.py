from pathlib import Path

from .base import register

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
