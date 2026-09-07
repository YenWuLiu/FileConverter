import json
import sys

from file_converter.__main__ import main


def test_no_args_prints_help(capsys):
    assert main([]) == 0
    assert "convert" in capsys.readouterr().out


def test_frozen_no_args_launches_gui(monkeypatch):
    import file_converter.gui as gui

    called = []
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(gui, "run", lambda: called.append(True))
    assert main([]) == 0
    assert called


def test_list_runs(capsys):
    assert main(["list"]) == 0
    out = capsys.readouterr().out
    assert "docx" in out and "png" in out


def test_console_encoding_never_crashes(monkeypatch):
    import io

    # 非中文 Windows（如 cp1252 控制台）无法编码中文输出，不应崩溃
    monkeypatch.setattr(sys, "stdout", io.TextIOWrapper(io.BytesIO(), encoding="cp1252"))
    assert main(["list"]) == 0


def test_convert_implicit_subcommand(tmp_path, capsys):
    src = tmp_path / "a.csv"
    src.write_text("k,v\n1,2\n", encoding="utf-8")
    assert main([str(src), "-t", "json", "-o", str(tmp_path)]) == 0
    assert json.loads((tmp_path / "a.json").read_text(encoding="utf-8")) == [{"k": "1", "v": "2"}]


def test_convert_failure_continues(tmp_path, capsys):
    good = tmp_path / "g.csv"
    good.write_text("a\n1\n", encoding="utf-8")
    bad = tmp_path / "b.csv"
    bad.write_text("a\n1\n", encoding="utf-8")
    bad.rename(tmp_path / "b.nope")  # 不支持的扩展名
    rc = main([str(good), str(tmp_path / "b.nope"), "-t", "json", "-o", str(tmp_path)])
    assert rc == 1  # 有失败
    assert (tmp_path / "g.json").exists()  # 但成功的照常产出
    assert "失败" in capsys.readouterr().out


def test_merge_split_rotate(tmp_path):
    import fitz
    from pypdf import PdfReader

    def mk(name, n):
        doc = fitz.open()
        for i in range(n):
            page = doc.new_page()
            page.insert_text((72, 72), f"{name}{i}")
        doc.save(str(tmp_path / name))
        doc.close()

    mk("m1.pdf", 2)
    mk("m2.pdf", 1)
    merged = tmp_path / "merged.pdf"
    assert main(["merge", str(tmp_path / "m1.pdf"), str(tmp_path / "m2.pdf"), "-o", str(merged)]) == 0
    assert len(PdfReader(str(merged)).pages) == 3

    outdir = tmp_path / "parts"
    assert main(["split", str(merged), "--pages", "1-2,3", "-o", str(outdir)]) == 0
    assert len(list(outdir.glob("*.pdf"))) == 2

    rot = tmp_path / "rot.pdf"
    assert main(["rotate", str(tmp_path / "m2.pdf"), "--angle", "180", "-o", str(rot)]) == 0
    assert PdfReader(str(rot)).pages[0].rotation == 180


def test_merge_no_overwrite(tmp_path):
    import fitz

    def mk(name):
        doc = fitz.open()
        doc.new_page()
        doc.save(str(tmp_path / name))
        doc.close()

    mk("a.pdf")
    mk("b.pdf")
    out = tmp_path / "out.pdf"
    out.write_bytes(b"keep me")
    assert main(["merge", str(tmp_path / "a.pdf"), str(tmp_path / "b.pdf"), "-o", str(out)]) == 0
    assert out.read_bytes() == b"keep me"
    assert (tmp_path / "out (1).pdf").exists()


def test_rotate_no_overwrite(tmp_path):
    import fitz

    doc = fitz.open()
    doc.new_page()
    doc.save(str(tmp_path / "r.pdf"))
    doc.close()

    out = tmp_path / "rot.pdf"
    out.write_bytes(b"keep me")
    assert main(["rotate", str(tmp_path / "r.pdf"), "--angle", "90", "-o", str(out)]) == 0
    assert out.read_bytes() == b"keep me"
    assert (tmp_path / "rot (1).pdf").exists()
