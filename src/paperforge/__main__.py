# -*- coding: utf-8 -*-
"""程序入口。

* 无参数（或双击 exe）→ 启动图形界面；
* 唯一参数是题库文件 → 启动图形界面并自动导入该题库（支持拖拽文件到 exe 上）；
* 其他参数 → 命令行模式（打包为 GUI 程序时会自动附着到父控制台）。
"""
from __future__ import annotations

import sys
from pathlib import Path

from .bank_io import SUPPORTED_SUFFIXES


def main(argv: "list[str] | None" = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)

    if args and args[0] == "--gui-selftest":
        import argparse

        from .cli import attach_parent_console

        parser = argparse.ArgumentParser(prog="PaperForge --gui-selftest",
                                         description="界面无头自检（发布验收用）")
        parser.add_argument("--gui-selftest", action="store_true")
        parser.add_argument("--bank", required=True)
        parser.add_argument("--out", required=True)
        parser.add_argument("--result", default="")
        parser.add_argument("--single", type=int, default=20)
        parser.add_argument("--multiple", type=int, default=10)
        parser.add_argument("--judge", type=int, default=10)
        parser.add_argument("--seed", type=int, default=12345)
        ns = parser.parse_args(args)
        attach_parent_console()
        from .gui import selftest

        code = selftest(ns.bank, ns.out, ns.result, ns.single, ns.multiple, ns.judge, ns.seed)
        try:
            sys.stdout.flush()
        except Exception:
            pass
        return code

    if args and args[0] in ("--gui",):
        args = args[1:]
        from .gui import run as run_gui

        return run_gui(args[0] if args else None)

    if len(args) == 1 and not args[0].startswith("-"):
        candidate = Path(args[0])
        if candidate.exists() and candidate.suffix.lower() in SUPPORTED_SUFFIXES:
            from .gui import run as run_gui

            return run_gui(str(candidate))

    if args:
        from .cli import attach_parent_console
        from .cli import main as run_cli

        attach_parent_console()
        try:
            return run_cli(args)
        finally:
            try:
                sys.stdout.flush()
                sys.stderr.flush()
            except Exception:
                pass

    from .gui import run as run_gui

    return run_gui()


if __name__ == "__main__":
    raise SystemExit(main())
