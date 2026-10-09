# -*- coding: utf-8 -*-
"""抽题引擎（核心领域逻辑）。

覆盖式抽题算法
--------------
1. 按题型分别维护一个"牌堆"（题库洗牌后的顺序列表）；
2. 第 i 张卷子从牌堆中顺序切出第 i 块，长度等于该题型每卷配额 k；
3. 牌堆取尽后进入补题：先放入剩余题目，再从**本轮已出现过**的题目中随机补足，
   使每张卷子题量严格等于配额；
4. 一轮张数 N = max(⌈该题型题库量 / 该题型配额⌉)，因此一轮结束后
   每种题型的题库都被完整覆盖（证明见 docs/02-概要设计说明书.md §5.4）。

选项打乱
--------
单选/多选可分别开关；打乱后答案字母按新选项位置重映射。
题库答案中带「定」标记的题目（如 ``BC（定）``）永不打乱。
"""
from __future__ import annotations

import math
import random
from dataclasses import dataclass, field

from .models import OPTION_LABELS, Blueprint, GenReport, Paper, PaperItem, QType, Question

__all__ = ["GenerationError", "GenOptions", "generate", "min_round_length",
           "group_by_type", "build_paper_item", "shuffle_options_of"]


class GenerationError(Exception):
    """组卷参数或题库不满足生成条件。"""


@dataclass
class GenOptions:
    """生成参数。"""

    blueprint: Blueprint = field(default_factory=Blueprint)
    round_length: "int | None" = None      # None/0 表示按覆盖性自动计算
    rounds: int = 1
    shuffle_single: bool = True
    shuffle_multiple: bool = True
    shuffle_items_in_paper: bool = True
    seed: "int | None" = None


# ---------------------------------------------------------------- 工具函数
def group_by_type(questions: list[Question]) -> dict[QType, list[Question]]:
    """按题型分组，保持题库原顺序。"""
    groups: dict[QType, list[Question]] = {QType.SINGLE: [], QType.MULTIPLE: [], QType.JUDGE: []}
    for q in questions:
        groups[q.qtype].append(q)
    return groups


def min_round_length(questions: list[Question], blueprint: Blueprint) -> int:
    """计算覆盖题库全部题目所需的最少张数（轮长）。"""
    groups = group_by_type(questions)
    need = 0
    for qtype, quota in blueprint.items():
        if quota <= 0:
            continue
        need = max(need, math.ceil(len(groups[qtype]) / quota))
    return max(need, 1)


def _label_for(index: int) -> str:
    """0 -> A, 1 -> B, ..., 25 -> Z, 26 -> AA。"""
    if index < 26:
        return OPTION_LABELS[index]
    return f"{index + 1}"


def build_paper_item(question: Question, rng: random.Random, do_shuffle: bool,
                     reused: bool = False) -> PaperItem:
    """把题目渲染为试卷题目项；``do_shuffle`` 为真且题目允许时打乱选项并重映射答案。"""
    n = question.option_count
    allow = do_shuffle and n >= 2 and not question.fixed_order and question.qtype is not QType.JUDGE
    if allow:
        perm = list(range(n))
        rng.shuffle(perm)
        labels = [OPTION_LABELS[j] for j in range(n)]
        texts = [question.options[i] for i in perm]
        position = {orig: j for j, orig in enumerate(perm)}
        new_letters = sorted(
            OPTION_LABELS[position[OPTION_LABELS.index(ch)]] for ch in question.answer_letters
        )
        return PaperItem(question=question, labels=labels, texts=texts,
                         answer_display="".join(new_letters), shuffled=True, reused=reused)
    labels = [OPTION_LABELS[j] for j in range(n)]
    return PaperItem(question=question, labels=labels, texts=list(question.options),
                     answer_display=question.answer_display, shuffled=False, reused=reused)


def shuffle_options_of(item: PaperItem, rng: random.Random) -> PaperItem:
    """对已生成的题目项按同一规则重新打乱（供独立测试/重排使用）。"""
    return build_paper_item(item.question, rng, True, reused=item.reused)


# ---------------------------------------------------------------- 主算法
def generate(questions: list[Question], options: GenOptions) -> GenReport:
    """生成若干轮试卷，并做覆盖性自检。"""
    bank = list(questions)
    if not bank:
        raise GenerationError("题库为空，无法生成试卷。请先导入题库。")
    blueprint = options.blueprint
    if blueprint.total <= 0:
        raise GenerationError("每张试卷的题目数量为 0，请至少为一种题型设置数量。")

    groups = group_by_type(bank)
    active: list[tuple[QType, int]] = [(t, k) for t, k in blueprint.items() if k > 0]

    for qtype, quota in active:
        if not groups[qtype]:
            raise GenerationError(
                f"试卷要求「{qtype.value}」{quota} 题，但题库中没有{qtype.value}，无法生成。"
            )

    auto_round_length = min_round_length(bank, blueprint)
    round_length = options.round_length or auto_round_length
    if round_length <= 0:
        round_length = auto_round_length
    rounds = max(1, int(options.rounds))

    seed = options.seed
    if seed is None:
        seed = random.SystemRandom().randrange(1, 2 ** 31 - 1)
    rng = random.Random(seed)

    warnings: list[str] = []
    if round_length < auto_round_length:
        warnings.append(
            f"一轮设置为 {round_length} 张，少于覆盖题库所需的最少张数 {auto_round_length} 张，"
            f"因此无法覆盖题库全部题目。"
        )
    for qtype, quota in active:
        n = len(groups[qtype])
        if n < quota:
            warnings.append(
                f"题库中{qtype.value}仅 {n} 题，少于每卷 {quota} 题，同一张卷子内将出现重复题目。"
            )

    total_by_type = {qtype: len(groups[qtype]) for qtype, _ in active}
    per_round_seen: list[dict[QType, set[str]]] = []
    per_round_missing: list[dict[QType, list[str]]] = []

    papers: list[Paper] = []
    global_seq = 0

    for round_index in range(rounds):
        draws: dict[QType, list[list[Question]]] = {}
        reuse_flags: dict[QType, list[list[bool]]] = {}

        for qtype, quota in active:
            deck = list(groups[qtype])
            rng.shuffle(deck)
            papers_draw: list[list[Question]] = []
            flags_draw: list[list[bool]] = []
            for i in range(round_length):
                start = i * quota
                chunk = deck[start:start + quota]
                flags = [False] * len(chunk)
                if len(chunk) < quota:
                    # 牌堆取尽 → 从本轮已出现过的题目中随机补足
                    pool = deck[:start]
                    need = quota - len(chunk)
                    if pool:
                        take = min(need, len(pool))
                        picks = rng.sample(pool, take)
                        chunk = chunk + picks
                        flags = flags + [True] * len(picks)
                        need -= len(picks)
                    while need > 0:  # 题库总量小于配额，只能重复抽样
                        chunk = chunk + [rng.choice(deck)]
                        flags = flags + [True]
                        need -= 1
                papers_draw.append(chunk)
                flags_draw.append(flags)
            draws[qtype] = papers_draw
            reuse_flags[qtype] = flags_draw

        seen_this_round: dict[QType, set[str]] = {qtype: set() for qtype, _ in active}
        for i in range(round_length):
            global_seq += 1
            paper = Paper(seq=global_seq, round_seq=i + 1, label=_label_for(i), items={})
            for qtype, quota in active:
                items: list[PaperItem] = []
                do_shuffle = (qtype is QType.SINGLE and options.shuffle_single) or \
                             (qtype is QType.MULTIPLE and options.shuffle_multiple)
                for q, reused in zip(draws[qtype][i], reuse_flags[qtype][i]):
                    items.append(build_paper_item(q, rng, do_shuffle, reused=reused))
                if options.shuffle_items_in_paper:
                    rng.shuffle(items)
                paper.items[qtype] = items
                seen_this_round[qtype].update(it.question.qid for it in items)
            papers.append(paper)

        per_round_seen.append(seen_this_round)
        missing: dict[QType, list[str]] = {}
        for qtype, _ in active:
            missing[qtype] = sorted(q.qid for q in groups[qtype]
                                    if q.qid not in seen_this_round[qtype])
        per_round_missing.append(missing)

    # 覆盖性汇总：取各轮中最差的一轮
    coverage: dict[str, tuple[int, int]] = {}
    coverage_ok = True
    uncovered: list[str] = []
    for qtype, _ in active:
        best_worst = min(len(r[qtype]) for r in per_round_seen)
        coverage[qtype.key] = (best_worst, total_by_type[qtype])
        if best_worst < total_by_type[qtype]:
            coverage_ok = False
            worst_round = min(range(len(per_round_seen)), key=lambda k: len(per_round_seen[k][qtype]))
            uncovered.extend(per_round_missing[worst_round][qtype])

    report = GenReport(
        papers=papers,
        round_length=round_length,
        rounds=rounds,
        blueprint=blueprint,
        coverage=coverage,
        coverage_ok=coverage_ok,
        uncovered=sorted(uncovered),
        warnings=warnings,
        seed=seed,
        bank_total=len(bank),
    )
    return report
