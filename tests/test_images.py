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
