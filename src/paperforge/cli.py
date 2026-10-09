# -*- coding: utf-8 -*-
"""命令行入口：用于批量、脚本化组卷。

示例
----
    PaperForge.exe --cli --bank 题库.xlsx --single 20 --multiple 10 --judge 10 ^
        --out 输出目录 --formats docx,html --rounds 1

不带 ``--cli`` 参数时启动图形界面。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __app_name__, __version__
from .bank_io import (
    bank_stats,
    export_questions_to_xlsx,
    load_bank,
    write_template_xlsx,
)
from .exporter import ExportOptions, export_all
from .generator import GenerationError, GenOptions, generate, min_round_length
from .models import Blueprint


def attach_parent_console() -> None:
    """打包为 GUI 程序时，若从 cmd/PowerShell 调用则附着到父控制台以显示输出。"""
    if sys.stdout is not None and sys.stderr is not None:
        return
    if sys.platform != "win32":
        return
    try:
        import ctypes

        if ctypes.windll.kernel32.AttachConsole(-1):
            sys.stdout = open("CONOUT$", "w", encoding="utf-8", buffering=1, errors="replace")
            sys.stderr = open("CONOUT$", "w", encoding="utf-8", buffering=1, errors="replace")
    except Exception:
        pass


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="PaperForge", description=f"{__app_name__} v{__version__} 批量组卷命令行")
    p.add_argument("bank", nargs="?", help="题库文件（xlsx/csv/txt/json）")
    p.add_argument("--bank", dest="bank_opt", help="题库文件（等价于位置参数）")
    p.add_argument("--single", type=int, default=20, help="每卷单选题数量（默认 20）")
    p.add_argument("--multiple", type=int, default=10, help="每卷多选题数量（默认 10）")
    p.add_argument("--judge", type=int, default=10, help="每卷判断题数量（默认 10）")
    p.add_argument("--papers", type=int, default=0, help="一轮张数（默认自动按覆盖性计算）")
    p.add_argument("--rounds", type=int, default=1, help="生成轮数（默认 1）")
    p.add_argument("--out", default="out", help="输出目录（默认 ./out）")
    p.add_argument("--formats", default="docx", help="导出格式，逗号分隔：docx,html,txt")
    p.add_argument("--merge", action="store_true", help="全部试卷合并为一个文件")
    p.add_argument("--answers-in-paper", action="store_true", help="答案附在卷末（默认单独成文件）")
    p.add_argument("--analysis", action="store_true", help="答案中附带解析")
    p.add_argument("--no-score", action="store_true", help="不显示分值")
    p.add_argument("--title", default="试卷", help="试卷标题（默认「试卷」）")
    p.add_argument("--seed", type=int, default=None, help="随机种子（相同种子可复现）")
    p.add_argument("--no-shuffle-single", action="store_true", help="不打乱单选选项")
    p.add_argument("--no-shuffle-multiple", action="store_true", help="不打乱多选选项")
    p.add_argument("--selfcheck", action="store_true", help="生成后打印覆盖性自检详情")
    p.add_argument("--stats-only", action="store_true", help="只统计题库，不生成试卷")
    p.add_argument("--make-template", metavar="PATH", help="生成题库模板文件后退出")
    p.add_argument("--normalize-bank", metavar="PATH", help="把解析后的题库规范化导出为 xlsx 后退出")
    p.add_argument("--cli", action="store_true", help="强制使用命令行模式（打包后使用）")
    p.add_argument("--log", default="", help="把命令行输出同时写入指定文本文件（UTF-8）")
    p.add_argument("--version", action="version", version=f"{__app_name__} {__version__}")
    return p





class _Tee:
    """把输出同时写到多个流（控制台 + 日志文件），任一失败不影响另一个。"""

    def __init__(self, *streams) -> None:
        self.streams = [s for s in streams if s is not None]

    def write(self, text: str) -> int:
        for stream in self.streams:
            try:
                stream.write(text)
                stream.flush()
            except Exception:
                pass
        return len(text or "")

    def flush(self) -> None:
        for stream in self.streams:
            try:
                stream.flush()
            except Exception:
                pass


def main(argv: "list[str] | None" = None) -> int:
    args = build_parser().parse_args(argv)
    if getattr(args, "log", ""):
        try:
            log_path = Path(args.log)
            log_path.parent.mkdir(parents=True, exist_ok=True)
            handle = log_path.open("w", encoding="utf-8")
            sys.stdout = _Tee(sys.stdout, handle)
            sys.stderr = _Tee(sys.stderr, handle)
        except Exception as exc:
            print(f"[警告] 无法写入日志文件：{exc}")

    if args.make_template:
        path = write_template_xlsx(args.make_template)
        print(f"[OK] 题库模板已生成：{path}")
        return 0

    bank_path = args.bank_opt or args.bank
    if not bank_path:
        print("[错误] 未指定题库文件。用法：PaperForge.exe --cli --bank 题库.xlsx ...")
        return 2

    result = load_bank(bank_path)
    stats = bank_stats(result.questions)
    print(f"[题库] {bank_path}")
    print(f"       共 {stats['total']} 题：单选 {stats['single']}、多选 {stats['multiple']}、"
          f"判断 {stats['judge']}；含固定顺序题 {stats['fixed']} 道")
    if result.issues:
        print(f"[警告] 有 {len(result.issues)} 行未能导入：")
        for issue in result.issues[:10]:
            print(f"       - {issue}")
        if len(result.issues) > 10:
            print(f"       - ...（其余 {len(result.issues) - 10} 条省略）")
    if not result.questions:
        print("[错误] 题库中没有可用题目。")
        return 3

    if args.normalize_bank:
        path = export_questions_to_xlsx(result.questions, args.normalize_bank)
        print(f"[OK] 规范化题库已导出：{path}")
        return 0

    if args.stats_only:
        return 0

    blueprint = Blueprint(args.single, args.multiple, args.judge)
    if blueprint.total <= 0:
        print("[错误] 每卷题目数量为 0。")
        return 2
    auto_n = min_round_length(result.questions, blueprint)
    papers = args.papers if args.papers > 0 else auto_n
    print(f"[蓝图] 每卷 {blueprint.total} 题（{blueprint.describe()}）；"
          f"覆盖题库最少需 {auto_n} 张/轮；本次一轮 {papers} 张 × {max(1, args.rounds)} 轮")

    opts = GenOptions(
        blueprint=blueprint,
        round_length=papers,
        rounds=max(1, args.rounds),
        shuffle_single=not args.no_shuffle_single,
        shuffle_multiple=not args.no_shuffle_multiple,
        seed=args.seed,
    )
    try:
        report = generate(result.questions, opts)
    except GenerationError as exc:
        print(f"[错误] {exc}")
        return 4

    print(f"[生成] {report.summary()}（随机种子 {report.seed}）")
    for warn in report.warnings:
        print(f"[警告] {warn}")
    if args.selfcheck:
        for key, (seen, total) in report.coverage.items():
            name = {"single": "单选", "multiple": "多选", "judge": "判断"}.get(key, key)
            flag = "完整" if seen >= total else f"缺 {total - seen} 题"
            print(f"       - {name}：{seen}/{total} {flag}")

    # 校验每张卷子的题量
    bad = [(p.label, p.total) for p in report.papers if p.total != blueprint.total]
    if bad:
        print(f"[警告] 以下试卷题量与蓝图不一致：{bad}")

    export_opts = ExportOptions(
        out_dir=Path(args.out),
        formats=tuple(f.strip() for f in args.formats.split(",") if f.strip()),
        answers_separate=not args.answers_in_paper,
        include_analysis=args.analysis,
        merge=args.merge,
        show_score=not args.no_score,
        title=args.title,
        bank_name=Path(bank_path).stem,
        base_name=Path(bank_path).stem or "试卷",
    )
    exp = export_all(report, export_opts)
    for f in exp.files:
        print(f"[导出] {f}")
    for e in exp.errors:
        print(f"[错误] {e}")
    return 0 if exp.files else 5
