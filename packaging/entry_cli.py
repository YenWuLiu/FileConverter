import os
import sys

_LOG = (
    os.path.join(os.path.dirname(sys.executable), "boot_cli.log")
    if getattr(sys, "frozen", False)
    else os.devnull
)


def _t(msg):
    try:
        with open(_LOG, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except OSError:
        pass


_t("0 entry start")
from file_converter import deps

_t("1 deps")
from file_converter.converters import base

_t("2 base")
from file_converter.converters import pdf

_t("3 pdf")
from file_converter import registry

_t("4 registry")
from file_converter.__main__ import main

_t("5 __main__ imported")
sys.exit(main())
