# -*- coding: utf-8 -*-
"""图形界面（tkinter/ttk）。

界面按 5 步组织：题库 → 蓝图 → 生成 → 预览 → 导出。
所有异常都被捕获并以中文提示框展示，避免打包后闪退。
"""
from __future__ import annotations

import os
import subprocess
import sys
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import __app_name__, __version__
from .bank_io import (bank_stats, export_questions_to_xlsx, load_bank,
                      write_template_xlsx)
from .config import AppConfig
from .exporter import ExportOptions, export_all, paper_to_text, safe_filename
from .generator import GenOptions, GenerationError, generate, min_round_length
from .models import Blueprint, GenReport, QType

HELP_TEXT = """【抽题匠 PaperForge 使用说明】

一、准备题库
  支持 .xlsx / .xlsm / .csv / .txt / .json / .docx / .docm / .pdf。\n  推荐用 Excel（xlsx）；Word 支持表格型与段落文本型；PDF 取文本层（扫描件无效）。\n  列头为：
  题型 | 题目标题 | 选项A | 选项B | 选项C | 选项D | … | 解析 | 答案
  · 题型填：单选题 / 多选题 / 判断题
  · 判断题不需要选项
  · 多选题答案写成字母组合，如 ABCD
  · 若某题选项顺序不可打乱，答案后加“（定）”，例如：BC（定）
  菜单「文件 → 生成题库模板」可得到一个可直接填写的示例文件。

二、设置试卷蓝图
  分别填写每张试卷的单选题、多选题、判断题数量（不需要的题型填 0）。
  例：20 / 10 / 10 → 每卷 40 题。

三、一轮张数与覆盖
  软件自动计算“覆盖题库全部题目所需的最少张数”：
      轮长 N = max( ⌈该题型题库量 ÷ 该题型每卷数量⌉ )
  例：单选 101 题、每卷 20 题 → 需要 6 张；多选 32 题、每卷 10 题 → 需要 4 张；
      判断 30 题、每卷 10 题 → 需要 3 张；因此一轮 = 6 张，6 张卷子考完题库所有题目。
  当某题型题目取尽时，软件会从本轮已出现过的题目中随机抽题补足，
  保证每张卷子的题目数量与蓝图完全一致。

四、选项打乱
  可分别勾选“打乱单选选项顺序”“打乱多选选项顺序”，防止抄袭。
  答案会自动跟随选项位置重新映射；标注“（定）”的题目永不打乱。

五、导出
  可导出 Word(.docx)、可打印网页(.html)、文本(.txt)。
  答案可单独成文件，也可附在卷末；可附带解析。
  网页格式在浏览器中按 Ctrl+P 即可打印或另存为 PDF。

六、命令行（可选）
  PaperForge.exe --cli --bank 题库.xlsx --single 20 --multiple 10 --judge 10 --out 输出目录
  详细参数：PaperForge.exe --cli --help
"""


class PaperForgeApp(tk.Tk):
    """主窗口。"""

    def __init__(self, initial_bank: str | None = None) -> None:
        super().__init__()
        self.cfg = AppConfig.load()
        self.bank = None          # BankLoadResult
        self.report: GenReport | None = None
        self._preview_index = 0

        self.title(f"{__app_name__} v{__version__} —— 题库随机抽题组卷工具")
        self.geometry("1160x820")
        self.minsize(1000, 700)

        self._init_vars()
        self._build_menu()
        self._build_layout()
        self._refresh_auto_round_length()

        bank = initial_bank or self.cfg.bank_path
        if bank and Path(bank).exists():
            self.after(120, lambda: self._load_bank(bank))

    # ------------------------------------------------------------ 变量
    def _init_vars(self) -> None:
        c = self.cfg
        self.bank_path_var = tk.StringVar(value=c.bank_path)
        self.stats_var = tk.StringVar(value="尚未导入题库")
        self.single_var = tk.StringVar(value=str(c.single_per_paper))
        self.multiple_var = tk.StringVar(value=str(c.multiple_per_paper))
        self.judge_var = tk.StringVar(value=str(c.judge_per_paper))
        self.total_var = tk.StringVar(value="")
        self.auto_round_var = tk.BooleanVar(value=bool(c.auto_round_length))
        self.manual_round_var = tk.StringVar(value=str(c.round_length or 0))
        self.rounds_var = tk.StringVar(value=str(c.rounds))
        self.shuffle_single_var = tk.BooleanVar(value=bool(c.shuffle_single))
        self.shuffle_multiple_var = tk.BooleanVar(value=bool(c.shuffle_multiple))
        self.seed_var = tk.StringVar(value=c.seed_text)
        self.auto_hint_var = tk.StringVar(value="")
        self.coverage_var = tk.StringVar(value="尚未生成试卷")
        self.paper_pick_var = tk.StringVar(value="")
        self.show_answer_var = tk.BooleanVar(value=False)
        self.out_dir_var = tk.StringVar(value=c.out_dir or str(Path.home() / "Desktop" / "抽题匠输出"))
        self.fmt_docx_var = tk.BooleanVar(value="docx" in c.formats)
        self.fmt_html_var = tk.BooleanVar(value="html" in c.formats)
        self.fmt_txt_var = tk.BooleanVar(value="txt" in c.formats)
        self.answers_separate_var = tk.BooleanVar(value=bool(c.answers_separate))
        self.analysis_var = tk.BooleanVar(value=bool(c.include_analysis))
        self.merge_var = tk.BooleanVar(value=bool(c.merge))
        self.show_score_var = tk.BooleanVar(value=bool(c.show_score))
        self.title_var = tk.StringVar(value=c.paper_title)
        self.status_var = tk.StringVar(value="就绪")

        for var in (self.single_var, self.multiple_var, self.judge_var):
            var.trace_add("write", lambda *_: self._refresh_auto_round_length())

    # ------------------------------------------------------------ 菜单
    def _build_menu(self) -> None:
        menubar = tk.Menu(self)
        m_file = tk.Menu(menubar, tearoff=0)
        m_file.add_command(label="导入题库…", accelerator="Ctrl+O", command=self._choose_bank)
        m_file.add_command(label="生成题库模板…", command=self._make_template)
        m_file.add_command(label="把当前题库规范化导出…", command=self._normalize_bank)
        m_file.add_separator()
        m_file.add_command(label="打开导出目录", command=self._open_out_dir)
        m_file.add_separator()
        m_file.add_command(label="退出", command=self.destroy)
        menubar.add_cascade(label="文件", menu=m_file)

        m_gen = tk.Menu(menubar, tearoff=0)
        m_gen.add_command(label="生成试卷", accelerator="F5", command=self._generate)
        m_gen.add_command(label="导出", accelerator="F6", command=self._export)
        menubar.add_cascade(label="生成", menu=m_gen)

        m_help = tk.Menu(menubar, tearoff=0)
        m_help.add_command(label="使用说明", accelerator="F1", command=self._show_help)
        m_help.add_command(label="关于", command=self._show_about)
        menubar.add_cascade(label="帮助", menu=m_help)

        self.config(menu=menubar)
        self.bind("<Control-o>", lambda e: self._choose_bank())
        self.bind("<F5>", lambda e: self._generate())
        self.bind("<F6>", lambda e: self._export())
        self.bind("<F1>", lambda e: self._show_help())

    # ------------------------------------------------------------ 布局
    def _build_layout(self) -> None:
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(3, weight=1)

        self._build_bank_box(root).grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self._build_plan_box(root).grid(row=1, column=0, sticky="ew", pady=(0, 8))
        self._build_generate_box(root).grid(row=2, column=0, sticky="ew", pady=(0, 8))
        self._build_preview_box(root).grid(row=3, column=0, sticky="nsew", pady=(0, 8))
        self._build_export_box(root).grid(row=4, column=0, sticky="ew")

        status = ttk.Frame(self, relief="sunken", padding=(8, 3))
        status.pack(fill="x", side="bottom")
        ttk.Label(status, textvariable=self.status_var, anchor="w").pack(fill="x")

    def _build_bank_box(self, parent) -> ttk.Widget:
        box = ttk.LabelFrame(parent, text=" ① 题库 ", padding=8)
        box.columnconfigure(1, weight=1)
        ttk.Label(box, text="题库文件：").grid(row=0, column=0, sticky="w")
        ttk.Entry(box, textvariable=self.bank_path_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(box, text="浏览…", command=self._choose_bank).grid(row=0, column=2, padx=2)
        ttk.Button(box, text="导入", command=lambda: self._load_bank(self.bank_path_var.get())
                   ).grid(row=0, column=3, padx=2)
        ttk.Label(box, textvariable=self.stats_var, foreground="#0a5").grid(
            row=1, column=0, columnspan=4, sticky="w", pady=(6, 0))
        return box

    def _build_plan_box(self, parent) -> ttk.Widget:
        box = ttk.LabelFrame(parent, text=" ② 试卷蓝图（每张试卷的题目数量） ", padding=8)
        for col in range(8):
            box.columnconfigure(col, weight=0)
        box.columnconfigure(7, weight=1)

        ttk.Label(box, text="单选题/卷").grid(row=0, column=0, sticky="e")
        ttk.Spinbox(box, from_=0, to=999, width=6, textvariable=self.single_var).grid(row=0, column=1, padx=(4, 12))
        ttk.Label(box, text="多选题/卷").grid(row=0, column=2, sticky="e")
        ttk.Spinbox(box, from_=0, to=999, width=6, textvariable=self.multiple_var).grid(row=0, column=3, padx=(4, 12))
        ttk.Label(box, text="判断题/卷").grid(row=0, column=4, sticky="e")
        ttk.Spinbox(box, from_=0, to=999, width=6, textvariable=self.judge_var).grid(row=0, column=5, padx=(4, 12))
        ttk.Label(box, textvariable=self.total_var, foreground="#036").grid(row=0, column=6, sticky="w")

        ttk.Checkbutton(box, text="打乱单选选项顺序（答案自动重映射）",
                        variable=self.shuffle_single_var).grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))
        ttk.Checkbutton(box, text="打乱多选选项顺序（标注「定」的题目除外）",
                        variable=self.shuffle_multiple_var).grid(row=1, column=3, columnspan=5, sticky="w", pady=(8, 0))

        ttk.Label(box, text="一轮张数").grid(row=2, column=0, sticky="e", pady=(8, 0))
        ttk.Radiobutton(box, text="自动（覆盖题库所需）", variable=self.auto_round_var, value=True,
                        command=self._refresh_auto_round_length).grid(row=2, column=1, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Radiobutton(box, text="手动", variable=self.auto_round_var, value=False,
                        command=self._refresh_auto_round_length).grid(row=2, column=3, sticky="w", pady=(8, 0))
        ttk.Spinbox(box, from_=1, to=999, width=6, textvariable=self.manual_round_var).grid(
            row=2, column=4, sticky="w", pady=(8, 0))
        ttk.Label(box, text="轮数").grid(row=2, column=5, sticky="e", pady=(8, 0))
        ttk.Spinbox(box, from_=1, to=99, width=6, textvariable=self.rounds_var).grid(
            row=2, column=6, sticky="w", pady=(8, 0))

        ttk.Label(box, text="随机种子（留空=每次不同）").grid(row=3, column=0, columnspan=2, sticky="e", pady=(8, 0))
        ttk.Entry(box, textvariable=self.seed_var, width=16).grid(row=3, column=2, sticky="w", pady=(8, 0))
        ttk.Label(box, textvariable=self.auto_hint_var, foreground="#a60").grid(
            row=3, column=3, columnspan=5, sticky="w", pady=(8, 0))
        return box

    def _build_generate_box(self, parent) -> ttk.Widget:
        box = ttk.LabelFrame(parent, text=" ③ 生成与覆盖校验 ", padding=8)
        box.columnconfigure(1, weight=1)
        ttk.Button(box, text="生成试卷（F5）", command=self._generate).grid(row=0, column=0)
        ttk.Label(box, textvariable=self.coverage_var, wraplength=900, justify="left").grid(
            row=0, column=1, sticky="w", padx=12)
        return box

    def _build_preview_box(self, parent) -> ttk.Widget:
        box = ttk.LabelFrame(parent, text=" ④ 预览 ", padding=8)
        box.columnconfigure(0, weight=1)
        box.rowconfigure(1, weight=1)
        top = ttk.Frame(box)
        top.grid(row=0, column=0, sticky="ew")
        ttk.Label(top, text="试卷：").pack(side="left")
        self.paper_combo = ttk.Combobox(top, textvariable=self.paper_pick_var, state="readonly", width=44)
        self.paper_combo.pack(side="left", padx=6)
        self.paper_combo.bind("<<ComboboxSelected>>", lambda e: self._render_preview())
        ttk.Checkbutton(top, text="显示答案", variable=self.show_answer_var,
                        command=self._render_preview).pack(side="left", padx=12)
        ttk.Label(top, text="（预览即导出内容；因题量较大，导出以文件为准）",
                  foreground="#888").pack(side="left")

        wrap = ttk.Frame(box)
        wrap.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        self.preview_text = tk.Text(wrap, wrap="word", font=("Microsoft YaHei", 10), undo=False)
        self.preview_text.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.preview_text.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.preview_text.configure(yscrollcommand=sb.set)
        self.preview_text.insert("1.0", "点击「生成试卷」后在此预览。")
        return box

    def _build_export_box(self, parent) -> ttk.Widget:
        box = ttk.LabelFrame(parent, text=" ⑤ 导出 ", padding=8)
        box.columnconfigure(1, weight=1)
        ttk.Label(box, text="导出目录：").grid(row=0, column=0, sticky="w")
        ttk.Entry(box, textvariable=self.out_dir_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(box, text="浏览…", command=self._choose_out_dir).grid(row=0, column=2)
        ttk.Button(box, text="打开", command=self._open_out_dir).grid(row=0, column=3, padx=(4, 0))

        opts = ttk.Frame(box)
        opts.grid(row=1, column=0, columnspan=4, sticky="w", pady=(8, 0))
        ttk.Label(opts, text="格式：").pack(side="left")
        ttk.Checkbutton(opts, text="Word(.docx)", variable=self.fmt_docx_var).pack(side="left")
        ttk.Checkbutton(opts, text="可打印网页(.html)", variable=self.fmt_html_var).pack(side="left", padx=6)
        ttk.Checkbutton(opts, text="文本(.txt)", variable=self.fmt_txt_var).pack(side="left")
        ttk.Checkbutton(opts, text="答案单独成文件", variable=self.answers_separate_var).pack(side="left", padx=(18, 0))
        ttk.Checkbutton(opts, text="答案含解析", variable=self.analysis_var).pack(side="left", padx=6)
        ttk.Checkbutton(opts, text="全部合并为一个文件", variable=self.merge_var).pack(side="left", padx=(18, 0))
        ttk.Checkbutton(opts, text="显示分值", variable=self.show_score_var).pack(side="left", padx=6)

        opts2 = ttk.Frame(box)
        opts2.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(6, 0))
        ttk.Label(opts2, text="试卷标题：").pack(side="left")
        ttk.Entry(opts2, textvariable=self.title_var, width=24).pack(side="left")
        ttk.Button(opts2, text="导出（F6）", command=self._export).pack(side="right")
        return box

    # ------------------------------------------------------------ 逻辑
    def _int_of(self, var: tk.StringVar, default: int = 0) -> int:
        try:
            return int(str(var.get()).strip())
        except (ValueError, TypeError):
            return default

    def _blueprint(self) -> Blueprint:
        return Blueprint(max(0, self._int_of(self.single_var)),
                         max(0, self._int_of(self.multiple_var)),
                         max(0, self._int_of(self.judge_var)))

    def _refresh_auto_round_length(self) -> None:
        bp = self._blueprint()
        self.total_var.set(f"→ 每卷合计 {bp.total} 题")
        if not self.bank or not self.bank.questions or bp.total <= 0:
            self.auto_hint_var.set("导入题库后自动计算覆盖所需张数")
            return
        try:
            n = min_round_length(self.bank.questions, bp)
        except Exception:
            n = 0
        stats = bank_stats(self.bank.questions)
        detail = []
        if bp.single > 0:
            detail.append(f"单选 {stats['single']}÷{bp.single}={-(-stats['single'] // bp.single)}张")
        if bp.multiple > 0:
            detail.append(f"多选 {stats['multiple']}÷{bp.multiple}={-(-stats['multiple'] // bp.multiple)}张")
        if bp.judge > 0:
            detail.append(f"判断 {stats['judge']}÷{bp.judge}={-(-stats['judge'] // bp.judge)}张")
        self.auto_hint_var.set(f"覆盖题库全部题目需 {n} 张/轮（" + "，".join(detail) + "）")
        if not self.auto_round_var.get():
            self.manual_round_var.set(str(n))

    def _choose_bank(self) -> None:
        path = filedialog.askopenfilename(
            title="选择题库文件",
            filetypes=[("题库文件", "*.xlsx *.xlsm *.csv *.txt *.json *.docx *.docm *.pdf"),
                       ("Excel 题库", "*.xlsx *.xlsm"),
                       ("Word 题库", "*.docx *.docm"),
                       ("PDF 题库", "*.pdf"),
                       ("文本/JSON 题库", "*.csv *.txt *.json"),
                       ("所有文件", "*.*")])
        if path:
            self.bank_path_var.set(path)
            self._load_bank(path)

    def _load_bank(self, path: str) -> None:
        path = (path or "").strip().strip('"')
        if not path:
            messagebox.showwarning("提示", "请先选择题库文件。")
            return
        if not Path(path).exists():
            messagebox.showerror("题库不存在", f"找不到文件：\n{path}")
            return
        self.config(cursor="watch")
        self.update_idletasks()
        try:
            result = load_bank(path)
        except Exception as exc:
            self.config(cursor="")
            messagebox.showerror("导入失败", f"{exc}")
            self.status_var.set("题库导入失败")
            return
        self.config(cursor="")
        self.bank = result
        self.report = None
        stats = bank_stats(result.questions)
        self.stats_var.set(
            f"已导入 {stats['total']} 题：单选 {stats['single']}、多选 {stats['multiple']}、"
            f"判断 {stats['judge']}；含固定顺序题 {stats['fixed']} 道"
            + (f"　｜　跳过 {len(result.issues)} 行" if result.issues else ""))
        self.status_var.set(f"题库导入成功：{path}")
        self.cfg.bank_path = path
        self._refresh_auto_round_length()
        if result.issues:
            detail = "\n".join(str(i) for i in result.issues[:15])
            more = f"\n…… 其余 {len(result.issues) - 15} 条省略" if len(result.issues) > 15 else ""
            messagebox.showwarning("部分行未能导入",
                                   f"以下 {len(result.issues)} 行未进入题库（其余题目已正常导入）：\n\n{detail}{more}")
        if not result.questions:
            messagebox.showerror("题库为空", "没有解析到任何可用题目，请检查题库格式。")

    def _generate(self) -> None:
        if not self.bank or not self.bank.questions:
            messagebox.showwarning("提示", "请先导入题库。")
            return
        bp = self._blueprint()
        if bp.total <= 0:
            messagebox.showwarning("提示", "请至少为一种题型设置每卷数量。")
            return
        round_length = None
        if not self.auto_round_var.get():
            round_length = self._int_of(self.manual_round_var, 0) or None
        opts = GenOptions(
            blueprint=bp,
            round_length=round_length,
            rounds=max(1, self._int_of(self.rounds_var, 1)),
            shuffle_single=self.shuffle_single_var.get(),
            shuffle_multiple=self.shuffle_multiple_var.get(),
            seed=self._seed_value(),
        )
        self.config(cursor="watch")
        self.update_idletasks()
        try:
            report = generate(self.bank.questions, opts)
        except GenerationError as exc:
            self.config(cursor="")
            messagebox.showerror("无法生成试卷", str(exc))
            return
        except Exception as exc:
            self.config(cursor="")
            messagebox.showerror("生成失败", f"{exc}\n\n{traceback.format_exc(limit=3)}")
            return
        self.config(cursor="")
        self.report = report
        self._preview_index = 0
        names = [f"第 {p.seq} 张（{p.name}，共 {p.total} 题）" for p in report.papers]
        self.paper_combo["values"] = names
        self.paper_pick_var.set(names[0] if names else "")
        status = "✔ " + report.summary()
        if report.warnings:
            status += "　｜　⚠ " + "；".join(report.warnings)
        self.coverage_var.set(status)
        self.status_var.set(f"已生成 {report.paper_count} 张试卷（随机种子 {report.seed}）")
        self._render_preview()

    def _seed_value(self) -> "int | None":
        text = (self.seed_var.get() or "").strip()
        if not text:
            return None
        try:
            return int(text)
        except ValueError:
            messagebox.showwarning("提示", "随机种子必须是整数，已按“每次随机”处理。")
            return None

    def _export_options(self) -> ExportOptions:
        formats = []
        if self.fmt_docx_var.get():
            formats.append("docx")
        if self.fmt_html_var.get():
            formats.append("html")
        if self.fmt_txt_var.get():
            formats.append("txt")
        return ExportOptions(
            out_dir=Path(self.out_dir_var.get().strip() or "."),
            formats=tuple(formats),
            answers_separate=self.answers_separate_var.get(),
            include_analysis=self.analysis_var.get(),
            merge=self.merge_var.get(),
            show_score=self.show_score_var.get(),
            title=self.title_var.get().strip() or "试卷",
            bank_name=Path(self.bank_path_var.get()).stem if self.bank_path_var.get() else "",
            scores=self.cfg.scores(),
            base_name=(self.title_var.get().strip() or "试卷"),
        )

    def _render_preview(self) -> None:
        idx = self.paper_combo.current()
        if idx < 0:
            idx = 0
        self._preview_index = idx
        self.preview_text.delete("1.0", "end")
        if not self.report or not self.report.papers:
            self.preview_text.insert("1.0", "点击「生成试卷」后在此预览。")
            return
        paper = self.report.papers[idx]
        try:
            text = paper_to_text(paper, self._export_options(), idx,
                                 include_answers=self.show_answer_var.get(),
                                 include_analysis=self.analysis_var.get())
        except Exception as exc:
            text = f"预览失败：{exc}"
        self.preview_text.insert("1.0", text)
        self.preview_text.see("1.0")

    def _export(self) -> None:
        if not self.report or not self.report.papers:
            messagebox.showwarning("提示", "请先生成试卷。")
            return
        opts = self._export_options()
        if not opts.formats:
            messagebox.showwarning("提示", "请至少选择一种导出格式。")
            return
        self.config(cursor="watch")
        self.status_var.set("正在导出……")
        self.update_idletasks()
        try:
            result = export_all(self.report, opts)
        except Exception as exc:
            self.config(cursor="")
            messagebox.showerror("导出失败", f"{exc}\n\n{traceback.format_exc(limit=3)}")
            return
        self.config(cursor="")
        if result.errors and not result.files:
            messagebox.showerror("导出失败", "\n".join(result.errors))
            self.status_var.set("导出失败")
            return
        self.cfg.out_dir = str(opts.out_dir)
        lines = "\n".join(str(f) for f in result.files[:12])
        more = f"\n…… 共 {len(result.files)} 个文件" if len(result.files) > 12 else ""
        msg = f"导出完成，共 {len(result.files)} 个文件：\n\n{lines}{more}"
        if result.errors:
            msg += "\n\n部分失败：\n" + "\n".join(result.errors)
        self.status_var.set(f"导出完成：{opts.out_dir}")
        if messagebox.askyesno("导出完成", msg + "\n\n是否打开导出目录？"):
            self._open_out_dir()

    def _choose_out_dir(self) -> None:
        path = filedialog.askdirectory(title="选择导出目录")
        if path:
            self.out_dir_var.set(path)

    def _open_out_dir(self) -> None:
        path = Path(self.out_dir_var.get().strip() or ".")
        try:
            path.mkdir(parents=True, exist_ok=True)
            if sys.platform == "win32":
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            messagebox.showerror("无法打开", f"{path}\n{exc}")

    def _make_template(self) -> None:
        path = filedialog.asksaveasfilename(title="保存题库模板", defaultextension=".xlsx",
                                            initialfile="题库模板.xlsx",
                                            filetypes=[("Excel 工作簿", "*.xlsx")])
        if not path:
            return
        try:
            write_template_xlsx(path)
            messagebox.showinfo("完成", f"模板已生成：\n{path}\n\n请按模板列头填写题目；"
                                        "多选题答案顺序固定时在答案后加“（定）”。")
        except Exception as exc:
            messagebox.showerror("生成失败", str(exc))

    def _normalize_bank(self) -> None:
        if not self.bank or not self.bank.questions:
            messagebox.showwarning("提示", "请先导入题库。")
            return
        path = filedialog.asksaveasfilename(title="规范化导出题库", defaultextension=".xlsx",
                                            initialfile="题库_规范化.xlsx",
                                            filetypes=[("Excel 工作簿", "*.xlsx")])
        if not path:
            return
        try:
            export_questions_to_xlsx(self.bank.questions, path)
            messagebox.showinfo("完成", f"已导出 {len(self.bank.questions)} 题：\n{path}")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))

    def _show_help(self) -> None:
        win = tk.Toplevel(self)
        win.title("使用说明")
        win.geometry("820x640")
        txt = tk.Text(win, wrap="word", font=("Microsoft YaHei", 10), padx=12, pady=10)
        txt.pack(fill="both", expand=True)
        txt.insert("1.0", HELP_TEXT)
        txt.configure(state="disabled")

    def _show_about(self) -> None:
        messagebox.showinfo(
            "关于",
            f"{__app_name__}  v{__version__}\n\n"
            "从题库随机抽题、按覆盖轮次组成试卷的桌面工具。\n"
            "支持单选题 / 多选题 / 判断题；一轮内覆盖题库全部题目，\n"
            "题量不足时自动从本轮已出现的题目中随机补足。\n\n"
            "全部计算在本机完成，不联网、不上传任何数据。\n"
            "单文件 exe，可自由复制分享。")

    def destroy(self) -> None:  # 退出前保存配置
        try:
            c = self.cfg
            c.bank_path = self.bank_path_var.get()
            c.single_per_paper = self._int_of(self.single_var, 20)
            c.multiple_per_paper = self._int_of(self.multiple_var, 10)
            c.judge_per_paper = self._int_of(self.judge_var, 10)
            c.auto_round_length = bool(self.auto_round_var.get())
            c.round_length = self._int_of(self.manual_round_var, 0)
            c.rounds = self._int_of(self.rounds_var, 1)
            c.shuffle_single = bool(self.shuffle_single_var.get())
            c.shuffle_multiple = bool(self.shuffle_multiple_var.get())
            c.seed_text = self.seed_var.get()
            c.out_dir = self.out_dir_var.get()
            c.formats = [f for f, v in (("docx", self.fmt_docx_var), ("html", self.fmt_html_var),
                                        ("txt", self.fmt_txt_var)) if v.get()]
            c.answers_separate = bool(self.answers_separate_var.get())
            c.include_analysis = bool(self.analysis_var.get())
            c.merge = bool(self.merge_var.get())
            c.show_score = bool(self.show_score_var.get())
            c.paper_title = self.title_var.get()
            c.save()
        except Exception:
            pass
        super().destroy()


def run(initial_bank: str | None = None) -> int:
    """启动图形界面。"""
    try:
        app = PaperForgeApp(initial_bank)
    except Exception as exc:  # 极少见：tk 初始化失败
        print(f"界面启动失败：{exc}", file=sys.stderr)
        return 1
    app.mainloop()
    return 0


def selftest(bank_path: str, out_dir: str, result_path: str = "",
             single: int = 20, multiple: int = 10, judge: int = 10,
             seed: int = 12345) -> int:
    """无头自检：在真实 tkinter 界面对象上完整走一遍流程，输出 JSON 结果。

    用于发布前的自动化验收（含打包后的 exe），不进入 mainloop。
    """
    import json
    from tkinter import messagebox as mb

    captured: list[dict] = []
    originals: dict = {}

    def _patch(name: str, kind: str) -> None:
        originals[name] = getattr(mb, name)

        def _fake(*a, **k):
            captured.append({"kind": kind, "text": str(a[0]) if a else ""})
            return False if kind == "ask" else None

        setattr(mb, name, _fake)

    for _n, _k in (("showinfo", "info"), ("showwarning", "warn"),
                   ("showerror", "error"), ("askyesno", "ask")):
        _patch(_n, _k)

    result: dict = {"ok": False, "stage": "init"}
    app: "PaperForgeApp | None" = None
    try:
        app = PaperForgeApp(None)
        app.withdraw()
        app.update()
        result["tk"] = {
            "screen": f"{app.winfo_screenwidth()}x{app.winfo_screenheight()}",
            "window_created": bool(app.winfo_exists()),
            "geometry": app.winfo_geometry(),
            "title": app.title(),
        }
        result["stage"] = "construct"

        app.bank_path_var.set(str(bank_path))
        app._load_bank(str(bank_path))
        result["bank"] = {
            "total": len(app.bank.questions) if app.bank else 0,
            "text": app.stats_var.get(),
        }
        if not app.bank or not app.bank.questions:
            raise RuntimeError("题库加载失败")

        app.single_var.set(str(single))
        app.multiple_var.set(str(multiple))
        app.judge_var.set(str(judge))
        app.seed_var.set(str(seed))
        app.auto_round_var.set(True)
        app.rounds_var.set("1")
        app.shuffle_single_var.set(True)
        app.shuffle_multiple_var.set(True)
        app.update_idletasks()
        result["stage"] = "plan"
        result["plan_hint"] = app.auto_hint_var.get()
        result["per_paper_total"] = app.total_var.get()

        app._generate()
        if not app.report:
            raise RuntimeError("生成失败")
        report = app.report
        result["stage"] = "generate"
        result["report"] = {
            "papers": report.paper_count,
            "round_length": report.round_length,
            "coverage_ok": report.coverage_ok,
            "coverage": {k: list(v) for k, v in report.coverage.items()},
            "paper_totals": [p.total for p in report.papers],
            "seed": report.seed,
            "coverage_text": app.coverage_var.get(),
        }
        preview = app.preview_text.get("1.0", "end")
        result["preview_chars"] = len(preview)
        result["preview_head"] = preview[:60]
        result["stage"] = "preview"

        out = Path(out_dir)
        out.mkdir(parents=True, exist_ok=True)
        app.out_dir_var.set(str(out))
        app.fmt_docx_var.set(True)
        app.fmt_html_var.set(True)
        app.fmt_txt_var.set(True)
        app.answers_separate_var.set(True)
        app.analysis_var.set(True)
        app.merge_var.set(False)
        app._export()
        files = sorted(p.name for p in out.iterdir() if p.is_file())
        result["file_count"] = len(files)
        result["files_sample"] = files[:4]
        result["ok"] = bool(report.coverage_ok and files and preview.strip())
        result["stage"] = "done"
    except Exception as exc:
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["traceback"] = traceback.format_exc(limit=8)
    finally:
        result["messages"] = captured
        for _n, _f in originals.items():
            setattr(mb, _n, _f)
        try:
            if app is not None:
                app.destroy()
        except Exception:
            pass
        payload = json.dumps(result, ensure_ascii=False, indent=2)
        if result_path:
            try:
                Path(result_path).parent.mkdir(parents=True, exist_ok=True)
                Path(result_path).write_text(payload, encoding="utf-8")
            except Exception:
                pass
        summary = {
            "ok": result.get("ok"),
            "stage": result.get("stage"),
            "papers": result.get("report", {}).get("papers"),
            "files": result.get("file_count"),
            "coverage_ok": result.get("report", {}).get("coverage_ok"),
            "error": result.get("error", ""),
        }
        try:
            print("[gui-selftest] " + json.dumps(summary, ensure_ascii=False))
            if not result_path:
                print(payload)
        except Exception:
            pass
    return 0 if result.get("ok") else 1
