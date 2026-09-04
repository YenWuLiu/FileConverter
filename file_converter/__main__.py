import argparse
import sys
from pathlib import Path

from . import deps, registry
from .converters.base import ConvertError
from .converters.pdf import merge_pdfs, rotate_pdf, split_pdf

_SUBCOMMANDS = ("convert", "merge", "split", "rotate", "list")


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="file_converter", description="万能文件转换工具（CLI）")
    p.add_argument("--gui", action="store_true", help="启动图形界面")
    sub = p.add_subparsers(dest="cmd")

    c = sub.add_parser("convert", help="格式转换（默认子命令，可省略）")
    c.add_argument("files", nargs="+", help="输入文件（可多个）")
    c.add_argument("-t", "--to", required=True, help="目标格式，如 pdf / jpg / xlsx")
    c.add_argument("-o", "--outdir", help="输出目录（默认与源文件同目录）")
    c.add_argument("--src-encoding", help="源文本编码（默认自动探测）")
    c.add_argument("--dst-encoding", default="utf-8", help="txt 输出编码（默认 utf-8）")
    c.add_argument("--dpi", type=int, default=150, help="PDF 转图片的 DPI（默认 150）")

    m = sub.add_parser("merge", help="合并多个 PDF")
    m.add_argument("files", nargs="+", help="至少 2 个 PDF")
    m.add_argument("-o", "--output", required=True, help="输出 PDF 路径")

    s = sub.add_parser("split", help="按页码拆分 PDF")
    s.add_argument("file", help="输入 PDF")
    s.add_argument("--pages", required=True, help="页码范围，如 1-3,5")
    s.add_argument("-o", "--outdir", required=True, help="输出目录")

    r = sub.add_parser("rotate", help="旋转 PDF 页面")
    r.add_argument("file", help="输入 PDF")
    r.add_argument("--angle", type=int, default=90, choices=(90, 180, 270), help="旋转角度")
    r.add_argument("-o", "--output", required=True, help="输出 PDF 路径")

    sub.add_parser("list", help="列出全部支持的转换")
    return p


def _cmd_list() -> int:
    print("支持的转换（[*] 表示依赖缺失，暂不可用）：\n")
    for srcs, dsts, label, ok in registry.matrix():
        mark = "" if ok else " [*]"
        print(f"  {', '.join(srcs):30s} → {', '.join(dsts):20s} {label}{mark}")
    if deps.office_available():
        print("\nOffice→PDF 引擎：MS Office (COM)")
    elif deps.libreoffice_path():
        print("\nOffice→PDF 引擎：LibreOffice")
    else:
        print("\nOffice→PDF 引擎：不可用（需安装 MS Office 或 LibreOffice）")
    return 0


def _cmd_convert(args) -> int:
    ok_count, fail_count = 0, 0
    for f in args.files:
        src = Path(f)
        if not src.exists():
            print(f"[失败] {f}: 文件不存在")
            fail_count += 1
            continue
        try:
            dst = None
            if args.outdir:
                dst = registry.unique_path(Path(args.outdir) / (src.stem + "." + args.to.lower().lstrip(".")))
            out = registry.convert(
                src, args.to, dst,
                src_encoding=args.src_encoding,
                dst_encoding=args.dst_encoding,
                dpi=args.dpi,
            )
            print(f"[成功] {src.name} → {out}")
            ok_count += 1
        except ConvertError as e:
            print(f"[失败] {src.name}: {e}")
            fail_count += 1
    print(f"\n完成：成功 {ok_count} 个，失败 {fail_count} 个")
    return 0 if fail_count == 0 else 1


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    parser = _build_parser()
    if argv and not argv[0].startswith("-") and argv[0] not in _SUBCOMMANDS:
        argv.insert(0, "convert")  # 省略子命令时按 convert 处理
    args = parser.parse_args(argv)

    if getattr(args, "gui", False) or args.cmd is None:
        # 打包成 exe（frozen）后无参数直接进入 GUI；源码运行无参数仍打印帮助
        if not getattr(args, "gui", False) and not getattr(sys, "frozen", False):
            parser.print_help()
            return 0
        from .gui import run

        run()
        return 0
    if args.cmd == "list":
        return _cmd_list()
    if args.cmd == "convert":
        return _cmd_convert(args)
    if args.cmd == "merge":
        try:
            out = registry.unique_path(Path(args.output))
            merge_pdfs([Path(f) for f in args.files], out)
            print(f"[成功] 已合并 → {out}")
            return 0
        except ConvertError as e:
            print(f"[失败] {e}")
            return 1
    if args.cmd == "split":
        try:
            outs = split_pdf(Path(args.file), Path(args.outdir), args.pages)
            for o in outs:
                print(f"[成功] {o}")
            return 0
        except ConvertError as e:
            print(f"[失败] {e}")
            return 1
    if args.cmd == "rotate":
        try:
            out = registry.unique_path(Path(args.output))
            rotate_pdf(Path(args.file), out, args.angle)
            print(f"[成功] 已旋转 {args.angle}° → {out}")
            return 0
        except ConvertError as e:
            print(f"[失败] {e}")
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
