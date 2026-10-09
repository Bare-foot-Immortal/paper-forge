# -*- coding: utf-8 -*-
"""抽题匠 PaperForge 测试套件（标准库 unittest，无第三方依赖）。

覆盖：答案解析、题型识别、选项打乱与答案重映射、轮长计算、
覆盖性、补题、随机种子复现、题库读写、导出回读、端到端合成题库。

运行：python -m unittest discover -s tests -v
"""
from __future__ import annotations

import json
import random
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from paperforge.bank_io import (  # noqa: E402
    bank_stats,
    build_sample_questions,
    export_questions_to_xlsx,
    load_bank,
    make_txt_sample,
    rows_to_questions,
    write_template_xlsx,
)
from paperforge.exporter import (  # noqa: E402
    ExportOptions,
    answer_to_text,
    export_all,
    paper_to_text,
    safe_filename,
)
from paperforge.generator import (  # noqa: E402
    GenerationError,
    GenOptions,
    build_paper_item,
    generate,
    min_round_length,
)
from paperforge.models import (  # noqa: E402
    OPTION_LABELS,
    AnswerParseError,
    Blueprint,
    QType,
    Question,
    parse_answer,
    parse_qtype,
)

FIXTURES = ROOT / "fixtures"
REAL_BANK = FIXTURES / "样例题库.xlsx"          # 合成夹具（tools/make_samples.py 生成）


def _ensure_fixtures() -> None:
    """夹具缺失时立即生成（必须在类定义/装饰器求值之前调用）。"""
    if REAL_BANK.exists():
        return
    import subprocess

    script = ROOT / "tools" / "make_samples.py"
    if script.exists():
        subprocess.run([sys.executable, str(script), "--fixtures-only"],
                       cwd=str(ROOT), check=False, capture_output=True)


_ensure_fixtures()




def make_question(index: int, qtype: QType = QType.SINGLE, n_options: int = 4,
                  fixed: bool = False, answer: "list[str] | None" = None) -> Question:
    if qtype is QType.JUDGE:
        return Question(qid=f"judge-{index:04d}", qtype=QType.JUDGE, stem=f"判断题 {index}",
                        options=[], judge_answer=(index % 2 == 0), fixed_order=fixed,
                        analysis="解析", src_row=index)
    letters = answer or (["A"] if qtype is QType.SINGLE else ["A", "B"])
    return Question(qid=f"{qtype.key}-{index:04d}", qtype=qtype, stem=f"题目 {index}（{qtype.short}）",
                    options=[f"选项{i}" for i in range(n_options)], answer_letters=list(letters),
                    fixed_order=fixed, analysis="解析", src_row=index)


# ================================================================ 模型层
class TestAnswerParsing(unittest.TestCase):
    def test_qtype_aliases(self):
        self.assertIs(parse_qtype("单选题"), QType.SINGLE)
        self.assertIs(parse_qtype("单选"), QType.SINGLE)
        self.assertIs(parse_qtype("多选题"), QType.MULTIPLE)
        self.assertIs(parse_qtype("judge"), QType.JUDGE)
        self.assertIs(parse_qtype("判断题"), QType.JUDGE)
        self.assertIsNone(parse_qtype("问答题"))

    def test_single_answer(self):
        letters, judge, fixed = parse_answer(QType.SINGLE, "A", 4)
        self.assertEqual(letters, ["A"])
        self.assertIsNone(judge)
        self.assertFalse(fixed)

    def test_multiple_answer_sorted(self):
        letters, _, _ = parse_answer(QType.MULTIPLE, "cba", 4)
        self.assertEqual(letters, ["A", "B", "C"])

    def test_multiple_answer_with_fixed_marker(self):
        # 需求示例：答案 BC（定）
        letters, judge, fixed = parse_answer(QType.MULTIPLE, "BC（定）", 4)
        self.assertEqual(letters, ["B", "C"])
        self.assertTrue(fixed)
        letters, _, fixed = parse_answer(QType.MULTIPLE, "bc(定)", 4)
        self.assertEqual(letters, ["B", "C"])
        self.assertTrue(fixed)

    def test_judge_answers(self):
        for raw, expected in [("正确", True), ("错误", False), ("对", True), ("错", False),
                              ("√", True), ("×", False), ("T", True), ("F", False),
                              ("是", True), ("否", False)]:
            letters, judge, _ = parse_answer(QType.JUDGE, raw, 0)
            self.assertEqual(letters, [])
            self.assertEqual(judge, expected, raw)

    def test_answer_out_of_range(self):
        with self.assertRaises(AnswerParseError):
            parse_answer(QType.SINGLE, "D", 3)

    def test_single_answer_multi_letters_rejected(self):
        with self.assertRaises(AnswerParseError):
            parse_answer(QType.SINGLE, "AB", 4)

    def test_empty_answer_rejected(self):
        with self.assertRaises(AnswerParseError):
            parse_answer(QType.SINGLE, "", 4)


# ================================================================ 抽题引擎
class TestShuffle(unittest.TestCase):
    def test_shuffle_keeps_answer_equivalent(self):
        """打乱后按新字母作答，命中与原答案完全相同的选项（非定题答案视为集合）。"""
        q = make_question(1, QType.MULTIPLE, n_options=4, answer=["A", "C"])
        rng = random.Random(7)
        for _ in range(200):
            item = build_paper_item(q, rng, True)
            picked = []
            for ch in item.answer_display:
                idx = OPTION_LABELS.index(ch)
                picked.append(item.texts[idx])
            expected = [q.options[OPTION_LABELS.index(c)] for c in q.answer_letters]
            self.assertEqual(sorted(picked), sorted(expected))
            self.assertEqual(len(picked), len(expected))   # 无重复、无遗漏
            self.assertTrue(item.shuffled)

    def test_shuffle_single_answer_points_to_same_option(self):
        q = make_question(9, QType.SINGLE, n_options=5, answer=["D"])
        rng = random.Random(99)
        expected_text = q.options[3]
        for _ in range(200):
            item = build_paper_item(q, rng, True)
            self.assertEqual(item.texts[OPTION_LABELS.index(item.answer_display)], expected_text)

    def test_fixed_order_answer_order_preserved(self):
        """「定」题不仅不打乱，答案顺序也保持题库原样（如 BC（定））。"""
        q = make_question(10, QType.MULTIPLE, n_options=4, fixed=True, answer=["B", "C"])
        rng = random.Random(5)
        for _ in range(50):
            item = build_paper_item(q, rng, True)
            self.assertEqual(item.answer_display, "BC")
            self.assertEqual(item.texts, q.options)

    def test_fixed_order_question_never_shuffled(self):
        q = make_question(2, QType.MULTIPLE, n_options=4, fixed=True, answer=["B", "C"])
        rng = random.Random(3)
        for _ in range(100):
            item = build_paper_item(q, rng, True)
            self.assertFalse(item.shuffled)
            self.assertEqual(item.texts, q.options)
            self.assertEqual(item.answer_display, "BC")

    def test_judge_not_shuffled(self):
        q = make_question(3, QType.JUDGE)
        item = build_paper_item(q, random.Random(1), True)
        self.assertFalse(item.shuffled)
        self.assertIn(item.answer_display, ("正确", "错误"))

    def test_shuffle_disabled(self):
        q = make_question(4, QType.SINGLE)
        item = build_paper_item(q, random.Random(1), False)
        self.assertFalse(item.shuffled)
        self.assertEqual(item.texts, q.options)
        self.assertEqual(item.answer_display, q.answer_display)


class TestRoundLength(unittest.TestCase):
    def test_min_round_length_matches_requirement(self):
        # 题库 101 单选 / 32 多选 / 30 判断，每卷 20/10/10 → 需 6 张
        questions = ([make_question(i, QType.SINGLE) for i in range(101)]
                     + [make_question(i, QType.MULTIPLE) for i in range(32)]
                     + [make_question(i, QType.JUDGE) for i in range(30)])
        self.assertEqual(min_round_length(questions, Blueprint(20, 10, 10)), 6)

    def test_round_length_ignores_zero_quota(self):
        questions = [make_question(i, QType.SINGLE) for i in range(10)]
        self.assertEqual(min_round_length(questions, Blueprint(3, 0, 0)), 4)


class TestGenerate(unittest.TestCase):
    def setUp(self):
        self.questions = ([make_question(i, QType.SINGLE) for i in range(101)]
                          + [make_question(i, QType.MULTIPLE) for i in range(32)]
                          + [make_question(i, QType.JUDGE) for i in range(30)])
        self.bp = Blueprint(20, 10, 10)

    def test_paper_size_equals_blueprint(self):
        report = generate(self.questions, GenOptions(blueprint=self.bp, seed=42))
        self.assertEqual(report.paper_count, 6)
        for paper in report.papers:
            self.assertEqual(paper.total, 40)
            self.assertEqual(paper.count_of(QType.SINGLE), 20)
            self.assertEqual(paper.count_of(QType.MULTIPLE), 10)
            self.assertEqual(paper.count_of(QType.JUDGE), 10)

    def test_full_coverage_in_one_round(self):
        report = generate(self.questions, GenOptions(blueprint=self.bp, seed=42))
        self.assertTrue(report.coverage_ok)
        self.assertEqual(report.uncovered, [])
        seen = {item.question.qid for p in report.papers for item in p.all_items()}
        self.assertEqual(seen, {q.qid for q in self.questions})
        self.assertEqual(report.coverage["single"], (101, 101))
        self.assertEqual(report.coverage["multiple"], (32, 32))
        self.assertEqual(report.coverage["judge"], (30, 30))

    def test_fill_up_with_used_questions(self):
        """多选 32 题、每卷 10 题：第 4 张应为 2 道新题 + 8 道补题。"""
        report = generate(self.questions, GenOptions(blueprint=self.bp, seed=42))
        paper4 = report.papers[3]
        multi = paper4.items_of(QType.MULTIPLE)
        self.assertEqual(len(multi), 10)
        reused = sum(1 for it in multi if it.reused)
        self.assertEqual(reused, 8)
        self.assertEqual(len({it.question.qid for it in multi}), 10)  # 卷内不重复

    def test_no_duplicate_in_single_paper(self):
        report = generate(self.questions, GenOptions(blueprint=self.bp, seed=1))
        for paper in report.papers:
            for qtype in (QType.SINGLE, QType.MULTIPLE, QType.JUDGE):
                items = paper.items_of(qtype)
                self.assertEqual(len(items), len({it.question.qid for it in items}))

    def test_seed_reproducible(self):
        a = generate(self.questions, GenOptions(blueprint=self.bp, seed=2024))
        b = generate(self.questions, GenOptions(blueprint=self.bp, seed=2024))
        self.assertEqual([[it.question.qid for it in p.all_items()] for p in a.papers],
                         [[it.question.qid for it in p.all_items()] for p in b.papers])
        self.assertEqual([[it.answer_display for it in p.all_items()] for p in a.papers],
                         [[it.answer_display for it in p.all_items()] for p in b.papers])

    def test_different_seed_differs(self):
        a = generate(self.questions, GenOptions(blueprint=self.bp, seed=1))
        b = generate(self.questions, GenOptions(blueprint=self.bp, seed=2))
        self.assertNotEqual([[it.question.qid for it in p.all_items()] for p in a.papers],
                            [[it.question.qid for it in p.all_items()] for p in b.papers])

    def test_multiple_rounds(self):
        report = generate(self.questions, GenOptions(blueprint=self.bp, seed=5, rounds=2))
        self.assertEqual(report.paper_count, 12)
        self.assertEqual(report.papers[0].label, "A")
        self.assertEqual(report.papers[6].label, "A")   # 第二轮重新从 A 卷开始
        self.assertTrue(report.coverage_ok)

    def test_manual_round_length_too_small_warns(self):
        report = generate(self.questions, GenOptions(blueprint=self.bp, seed=1, round_length=3))
        self.assertFalse(report.coverage_ok)
        self.assertTrue(report.uncovered)
        self.assertTrue(any("无法覆盖" in w for w in report.warnings))
        for paper in report.papers:
            self.assertEqual(paper.total, 40)

    def test_empty_bank_raises(self):
        with self.assertRaises(GenerationError):
            generate([], GenOptions(blueprint=Blueprint(1, 0, 0)))

    def test_zero_blueprint_raises(self):
        with self.assertRaises(GenerationError):
            generate(self.questions, GenOptions(blueprint=Blueprint(0, 0, 0)))

    def test_missing_type_raises(self):
        only_single = [make_question(i, QType.SINGLE) for i in range(5)]
        with self.assertRaises(GenerationError):
            generate(only_single, GenOptions(blueprint=Blueprint(2, 1, 0)))

    def test_quota_larger_than_bank(self):
        questions = [make_question(i, QType.SINGLE) for i in range(3)]
        report = generate(questions, GenOptions(blueprint=Blueprint(5, 0, 0), seed=1))
        self.assertEqual(report.paper_count, 1)
        self.assertEqual(report.papers[0].total, 5)
        self.assertTrue(any("重复" in w for w in report.warnings))

    def test_options_not_shuffled_when_disabled(self):
        report = generate(self.questions, GenOptions(blueprint=self.bp, seed=3,
                                                     shuffle_single=False,
                                                     shuffle_multiple=False))
        for paper in report.papers:
            for item in paper.all_items():
                if item.qtype is not QType.JUDGE:
                    self.assertFalse(item.shuffled)
                    self.assertEqual(item.answer_display, item.question.answer_display)


# ================================================================ 题库 IO
class TestBankIO(unittest.TestCase):
    def test_rows_to_questions_with_header(self):
        rows = [
            ["题型", "题目标题", "选项A", "选项B", "选项C", "选项D", "解析", "答案"],
            ["单选题", "题目一", "甲", "乙", "丙", "丁", "解析一", "B"],
            ["多选题", "题目二", "甲", "乙", "丙", "丁", "解析二", "AC（定）"],
            ["判断题", "题目三", "", "", "", "", "解析三", "正确"],
        ]
        questions, issues, header, cmap = rows_to_questions(rows)
        self.assertEqual(len(questions), 3)
        self.assertEqual(issues, [])
        self.assertEqual(questions[1].answer_letters, ["A", "C"])
        self.assertTrue(questions[1].fixed_order)
        self.assertTrue(questions[2].judge_answer)

    def test_bad_rows_collected(self):
        rows = [
            ["题型", "题目标题", "选项A", "选项B", "选项C", "选项D", "解析", "答案"],
            ["单选题", "好题", "甲", "乙", "丙", "丁", "解析", "A"],
            ["问答题", "坏题", "甲", "乙", "", "", "", "A"],
            ["单选题", "缺答案", "甲", "乙", "丙", "丁", "解析", ""],
            ["单选题", "答案越界", "甲", "乙", "丙", "丁", "解析", "F"],
            ["", "", "", "", "", "", "", ""],
        ]
        questions, issues, _, _ = rows_to_questions(rows)
        self.assertEqual(len(questions), 1)
        self.assertEqual(len(issues), 3)

    def test_load_sample_xlsx_csv_txt_json(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            sample = build_sample_questions()
            # xlsx
            xlsx = export_questions_to_xlsx(sample, tmp_path / "b.xlsx")
            res = load_bank(xlsx)
            self.assertEqual(len(res.questions), len(sample))
            self.assertEqual(res.issues, [])
            # csv
            csv_path = tmp_path / "b.csv"
            rows = [["题型", "题目标题", "选项A", "选项B", "选项C", "选项D", "解析", "答案"]]
            for q in sample:
                opts = (q.options + ["", "", "", ""])[:4]
                rows.append([q.qtype.value, q.stem, *opts, q.analysis,
                             q.answer_display + ("（定）" if q.fixed_order else "")])
            import csv as _csv

            with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
                _csv.writer(f).writerows(rows)
            res = load_bank(csv_path)
            self.assertEqual(len(res.questions), len(sample))
            # txt（制表符）
            txt = tmp_path / "b.txt"
            txt.write_text(make_txt_sample(), encoding="utf-8")
            res = load_bank(txt)
            self.assertGreaterEqual(len(res.questions), 6)
            # json
            js = tmp_path / "b.json"
            js.write_text(json.dumps({"questions": [
                {"type": "单选题", "stem": "JSON 题", "options": ["甲", "乙", "丙", "丁"],
                 "answer": "C", "analysis": "解析"},
                {"type": "判断题", "stem": "JSON 判断题", "answer": "错误"},
            ]}, ensure_ascii=False), encoding="utf-8")
            res = load_bank(js)
            self.assertEqual(len(res.questions), 2)
            self.assertEqual(res.questions[0].answer_letters, ["C"])
            self.assertFalse(res.questions[1].judge_answer)

    def test_template_and_stats(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = write_template_xlsx(Path(tmp) / "模板.xlsx")
            res = load_bank(path)
            self.assertGreaterEqual(len(res.questions), 1)
            stats = bank_stats(res.questions)
            self.assertEqual(stats["total"], len(res.questions))

    def test_missing_file(self):
        with self.assertRaises(FileNotFoundError):
            load_bank("不存在的题库.xlsx")

    def test_unsupported_suffix(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "x.doc"
            p.write_text("x", encoding="utf-8")
            with self.assertRaises(ValueError):
                load_bank(p)


# ================================================================ 导出
class TestExport(unittest.TestCase):
    def setUp(self):
        self.questions = ([make_question(i, QType.SINGLE) for i in range(12)]
                          + [make_question(i, QType.MULTIPLE, answer=["A", "B"]) for i in range(6)]
                          + [make_question(i, QType.JUDGE) for i in range(4)])
        self.bp = Blueprint(4, 2, 2)
        self.report = generate(self.questions, GenOptions(blueprint=self.bp, seed=11))

    def test_text_render_contains_structure(self):
        text = paper_to_text(self.report.papers[0], ExportOptions(), 0)
        self.assertIn("A卷", text)
        self.assertIn("单选题", text)
        self.assertIn("多选题", text)
        self.assertIn("判断题", text)
        ans = answer_to_text(self.report.papers[0], ExportOptions(), 0)
        self.assertIn("参考答案", ans)

    def test_export_all_formats(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp)
            opts = ExportOptions(out_dir=out, formats=("docx", "html", "txt"),
                                 answers_separate=True, include_analysis=True, title="试卷")
            result = export_all(self.report, opts)
            self.assertEqual(result.errors, [])
            # 6 张卷子：每张 试卷+答案 两个文件 × 3 种格式
            self.assertEqual(len(result.files), self.report.paper_count * 2 * 3)
            for f in result.files:
                self.assertTrue(f.exists() and f.stat().st_size > 0, f)
            docx_files = [f for f in result.files if f.suffix == ".docx"]
            self.assertEqual(len(docx_files), self.report.paper_count * 2)

    def test_docx_readback(self):
        from docx import Document

        with tempfile.TemporaryDirectory() as tmp:
            opts = ExportOptions(out_dir=Path(tmp), formats=("docx",), answers_separate=True)
            files = export_all(self.report, opts).files
            paper_file = [f for f in files if "答案" not in f.name][0]
            doc = Document(paper_file)
            text = "\n".join(p.text for p in doc.paragraphs)
            self.assertIn("A卷", text)
            self.assertIn("单选题", text)
            self.assertIn("姓名", text)
            answer_file = [f for f in files if "答案" in f.name][0]
            atext = "\n".join(p.text for p in Document(answer_file).paragraphs)
            self.assertIn("参考答案", atext)

    def test_merged_export(self):
        with tempfile.TemporaryDirectory() as tmp:
            opts = ExportOptions(out_dir=Path(tmp), formats=("docx", "html", "txt"),
                                 merge=True, answers_separate=True)
            files = export_all(self.report, opts).files
            self.assertEqual(len(files), 6)   # 3 种格式 × (试卷 + 答案)

    def test_answers_in_paper(self):
        with tempfile.TemporaryDirectory() as tmp:
            opts = ExportOptions(out_dir=Path(tmp), formats=("txt",), answers_separate=False)
            files = export_all(self.report, opts).files
            self.assertEqual(len(files), self.report.paper_count)
            content = files[0].read_text(encoding="utf-8-sig")
            self.assertIn("参考答案", content)

    def test_safe_filename(self):
        self.assertEqual(safe_filename('a\\b/c:d*e?f"g<h>i|j'), "a_b_c_d_e_f_g_h_i_j")
        self.assertEqual(safe_filename(""), "未命名")


# ================================================================ 端到端
@unittest.skipUnless(REAL_BANK.exists(), f"未找到合成题库：{REAL_BANK}")
class TestRealBankEndToEnd(unittest.TestCase):
    """使用桌面合成题库（52 题）做端到端验收。"""

    @classmethod
    def setUpClass(cls):
        cls.result = load_bank(REAL_BANK)

    def test_bank_fully_parsed(self):
        stats = bank_stats(self.result.questions)
        self.assertEqual(stats["total"], 52)
        self.assertEqual(stats["single"], 26)
        self.assertEqual(stats["multiple"], 14)
        self.assertEqual(stats["judge"], 12)
        self.assertEqual(stats["fixed"], 1)
        self.assertEqual(self.result.issues, [])

    def test_round_covers_every_question(self):
        bp = Blueprint(20, 10, 10)
        report = generate(self.result.questions, GenOptions(blueprint=bp, seed=2026))
        self.assertEqual(report.paper_count, 2)
        self.assertTrue(report.coverage_ok)
        self.assertEqual(report.uncovered, [])
        self.assertEqual(report.coverage,
                         {"single": (26, 26), "multiple": (14, 14), "judge": (12, 12)})
        for paper in report.papers:
            self.assertEqual(paper.total, 40)
        seen = {it.question.qid for p in report.papers for it in p.all_items()}
        self.assertEqual(len(seen), 52)

    def test_export_real_papers_to_docx(self):
        bp = Blueprint(20, 10, 10)
        report = generate(self.result.questions, GenOptions(blueprint=bp, seed=1))
        with tempfile.TemporaryDirectory() as tmp:
            opts = ExportOptions(out_dir=Path(tmp), formats=("docx", "html", "txt"),
                                 answers_separate=True, include_analysis=True,
                                 title="安全培训考试试卷")
            result = export_all(report, opts)
            self.assertEqual(result.errors, [])
            self.assertTrue(all(f.exists() for f in result.files))
            total_size = sum(f.stat().st_size for f in result.files)
            self.assertGreater(total_size, 10000)


@unittest.skipUnless(REAL_BANK.exists(), f"未找到合成题库：{REAL_BANK}")
class TestGuiSelftest(unittest.TestCase):
    """界面层无头自检：真实 tkinter 对象上跑完整流程（导入→生成→预览→导出）。"""

    def test_gui_flow(self):
        try:
            import tkinter
            root = tkinter.Tk()
            root.destroy()
        except Exception as exc:                       # 无图形环境则跳过
            self.skipTest(f"当前环境不支持 tkinter：{exc}")

        from paperforge.gui import selftest

        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "out"
            res_file = Path(tmp) / "result.json"
            try:
                code = selftest(str(REAL_BANK), str(out), str(res_file), 20, 10, 10, 12345)
            except Exception as exc:                   # 例如 CI 上 Tcl 初始化失败
                self.skipTest(f"当前环境无法运行界面自检：{exc}")
            self.assertEqual(code, 0, res_file.read_text(encoding="utf-8") if res_file.exists() else "")
            data = json.loads(res_file.read_text(encoding="utf-8"))
            self.assertTrue(data["tk"]["window_created"])
            self.assertEqual(data["bank"]["total"], 52)
            self.assertTrue(data["report"]["coverage_ok"])
            self.assertEqual(data["report"]["papers"], 2)
            self.assertEqual(data["report"]["paper_totals"], [40, 40])
            self.assertEqual(data["report"]["coverage"],
                             {"single": [26, 26], "multiple": [14, 14], "judge": [12, 12]})
            self.assertGreater(data["preview_chars"], 1000)
            self.assertEqual(data["file_count"], 12)
            # 全流程不应出现错误/警告弹窗；仅允许导出完成后的"是否打开目录"询问
            bad = [m for m in data["messages"] if m["kind"] in ("error", "warn")]
            self.assertEqual(bad, [])
            self.assertTrue(all("导出完成" in m["text"] for m in data["messages"]))


if __name__ == "__main__":
    unittest.main(verbosity=2)
