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
        page.insert_font(fontname="cjk", fontfile="C:/Windows/Fonts/simsun.ttc")
        page.insert_text((72, 72), t, fontname="cjk")
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
    registry.convert(pdf3, "png")
    first = pdf3.parent / "a_p1.png"
    assert first.exists()
    assert (pdf3.parent / "a_p3.png").exists()
    assert first.read_bytes()[:8] == b"\x89PNG\r\n\x1a\n"


def test_pdf_to_png_no_overwrite(pdf3):
    registry.convert(pdf3, "png")
    registry.convert(pdf3, "png")
    assert (pdf3.parent / "a_p1.png").exists()
    assert (pdf3.parent / "a_p1 (1).png").exists()


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
