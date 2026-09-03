import json

from file_converter.__main__ import main


def test_list_runs(capsys):
    assert main(["list"]) == 0
    out = capsys.readouterr().out
    assert "docx" in out and "png" in out


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
