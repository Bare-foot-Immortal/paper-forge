# -*- coding: utf-8 -*-
"""配置持久化：界面参数保存到用户目录，下次启动自动恢复。

保存位置：``%APPDATA%/PaperForge/config.json``（Windows）。
不写注册表、不写程序目录，保证 exe 可绿色分发。
"""
from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass, field
from pathlib import Path

from .models import Blueprint

__all__ = ["AppConfig", "config_dir", "config_path"]


def config_dir() -> Path:
    base = os.environ.get("APPDATA") or os.environ.get("XDG_CONFIG_HOME")
    if base:
        return Path(base) / "PaperForge"
    return Path.home() / ".paperforge"


def config_path() -> Path:
    return config_dir() / "config.json"


@dataclass
class AppConfig:
    """界面/命令行共享的默认参数。"""

    bank_path: str = ""
    single_per_paper: int = 20
    multiple_per_paper: int = 10
    judge_per_paper: int = 10
    auto_round_length: bool = True
    round_length: int = 0
    rounds: int = 1
    shuffle_single: bool = True
    shuffle_multiple: bool = True
    seed_text: str = ""
    out_dir: str = ""
    formats: list[str] = field(default_factory=lambda: ["docx"])
    answers_separate: bool = True
    include_analysis: bool = False
    merge: bool = False
    show_score: bool = True
    exam_minutes: int = 0
    paper_title: str = "试卷"
    base_name: str = "试卷"
    score_single: float = 1.0
    score_multiple: float = 2.0
    score_judge: float = 1.0

    # ---------------------------------------------------------- 便捷方法
    def blueprint(self) -> Blueprint:
        return Blueprint(int(self.single_per_paper), int(self.multiple_per_paper),
                         int(self.judge_per_paper))

    def scores(self) -> dict:
        return {"single": float(self.score_single), "multiple": float(self.score_multiple),
                "judge": float(self.score_judge)}

    def seed(self) -> "int | None":
        text = (self.seed_text or "").strip()
        if not text:
            return None
        try:
            return int(text)
        except ValueError:
            return None

    # ---------------------------------------------------------- 持久化
    def save(self, path: "Path | None" = None) -> Path:
        p = Path(path) if path else config_path()
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(json.dumps(asdict(self), ensure_ascii=False, indent=2), encoding="utf-8")
        return p

    @classmethod
    def load(cls, path: "Path | None" = None) -> "AppConfig":
        p = Path(path) if path else config_path()
        if not p.exists():
            return cls()
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return cls()
        cfg = cls()
        for key, value in (data or {}).items():
            if not hasattr(cfg, key):
                continue
            # 逐字段按默认值类型转换：配置文件被改坏时取默认值，避免"双击没反应"
            current = getattr(cfg, key)
            try:
                if isinstance(current, bool):
                    cfg_value = bool(value)
                elif isinstance(current, int):
                    cfg_value = int(value)
                elif isinstance(current, float):
                    cfg_value = float(value)
                elif isinstance(current, str):
                    cfg_value = str(value)
                elif isinstance(current, (list, tuple)):
                    if not isinstance(value, (list, tuple)):
                        continue
                    cfg_value = type(current)(value)
                elif isinstance(current, dict):
                    if not isinstance(value, dict):
                        continue
                    cfg_value = dict(value)
                else:
                    cfg_value = value
            except (TypeError, ValueError):
                continue
            setattr(cfg, key, cfg_value)
        return cfg
