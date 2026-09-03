import csv
import subprocess
from pathlib import Path

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

    registry.convert(p, "csv")
    assert (tmp_path / "m_一.csv").exists()
    assert (tmp_path / "m_二.csv").exists()


def test_xlsx_to_csv_multisheet_no_overwrite(tmp_path):
    import openpyxl

    wb = openpyxl.Workbook()
    wb.active.title = "一"
    wb.create_sheet("二")
    p = tmp_path / "m.xlsx"
    wb.save(str(p))
    registry.convert(p, "csv")
    registry.convert(p, "csv")
    assert (tmp_path / "m_一.csv").exists()
    assert (tmp_path / "m_一 (1).csv").exists()


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

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "演示标题"
    slide.placeholders[1].text = "要点一"
    p = tmp_path / "a.pptx"
    prs.save(str(p))

    out = registry.convert(p, "txt")
    text = out.read_text(encoding="utf-8")
    assert "演示标题" in text and "要点一" in text


def test_libreoffice_to_pdf_does_not_overwrite_existing(tmp_path, monkeypatch):
    import docx

    from file_converter import deps
    from file_converter.converters import documents

    d = docx.Document()
    d.add_paragraph("内容")
    src = tmp_path / "a.docx"
    d.save(str(src))

    sentinel = tmp_path / "a.pdf"
    sentinel.write_bytes(b"SENTINEL-ORIGINAL")

    monkeypatch.setattr(deps, "office_available", lambda: False)
    monkeypatch.setattr(deps, "libreoffice_path", lambda: "soffice")

    def fake_run(cmd, **kwargs):
        outdir = Path(cmd[cmd.index("--outdir") + 1])
        assert outdir != tmp_path
        (outdir / "a.pdf").write_bytes(b"FAKE-PDF")
        return subprocess.CompletedProcess(cmd, 0, "", "")

    monkeypatch.setattr(documents.subprocess, "run", fake_run)

    dst = registry.unique_path(tmp_path / "a.pdf")
    assert dst != sentinel
    out = registry.convert(src, "pdf", dst)
    assert out == dst
    assert sentinel.read_bytes() == b"SENTINEL-ORIGINAL"
    assert out.read_bytes() == b"FAKE-PDF"
