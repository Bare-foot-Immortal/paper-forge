# -*- coding: utf-8 -*-
"""生成软件图标 assets/PaperForge.ico（构建期使用，不随程序运行）。

用 Pillow + 系统中文字体绘制：蓝色圆角方块 + 白色试卷 + 对勾。
若系统缺少字体或 Pillow，则跳过（打包脚本会退化为无自定义图标）。
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "assets" / "PaperForge.ico"

FONT_CANDIDATES = [
    r"C:\Windows\Fonts\msyhbd.ttc",
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\simsun.ttc",
]


def rounded_rect(draw, box, radius, fill):
    x0, y0, x1, y1 = box
    draw.rounded_rectangle([x0, y0, x1, y1], radius=radius, fill=fill)


def main() -> int:
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        print("[跳过] 未安装 Pillow，将使用默认图标。")
        return 0

    font_path = next((p for p in FONT_CANDIDATES if Path(p).exists()), None)
    base = 256
    img = Image.new("RGBA", (base, base), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # 背景：深蓝渐变圆角方块
    for i in range(base):
        t = i / base
        color = (int(21 + 30 * t), int(87 + 60 * t), int(178 + 45 * t), 255)
        d.line([(0, i), (base, i)], fill=color)
    mask = Image.new("L", (base, base), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, base - 1, base - 1], radius=52, fill=255)
    img.putalpha(mask)

    # 白色试卷
    sheet = [58, 44, 198, 212]
    rounded_rect(d, sheet, 14, (255, 255, 255, 255))
    # 试卷上的文字行
    for i, y in enumerate(range(74, 150, 22)):
        d.rounded_rectangle([78, y, 178 - i * 12, y + 9], radius=4, fill=(150, 175, 205, 255))

    # 主标题字
    if font_path:
        try:
            font = ImageFont.truetype(font_path, 96)
            d.text((base / 2, 168), "题", font=font, fill=(25, 82, 165, 255), anchor="mm")
        except Exception:
            pass

    # 右下角对勾
    d.line([(150, 196), (176, 220), (232, 148)], fill=(28, 165, 90, 255), width=22, joint="curve")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    img.save(OUT, sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
    print(f"[OK] 图标已生成：{OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
