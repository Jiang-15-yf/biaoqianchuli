# -*- coding: utf-8 -*-
"""生成价格牌处理工具的 Logo（PNG + ICO）

设计：深色圆角徽章 + 白色吊牌（挂孔 + 两行价格信息）。
采用 8 倍超采样渲染后降采样，保证 16px 等小尺寸依然锐利。
"""
from __future__ import annotations

from pathlib import Path
from PIL import Image, ImageDraw

HERE = Path(__file__).parent.resolve()

BADGE_COLOR = (23, 24, 28, 255)      # #17181C，与 UI 主色 INK 一致
CARD_COLOR = (255, 255, 255, 255)    # 白色吊牌


def _round_rect(draw: ImageDraw.Draw, xy, radius, fill=None):
    """绘制实心圆角矩形"""
    x1, y1, x2, y2 = xy
    r = max(0, min(radius, (x2 - x1) // 2, (y2 - y1) // 2))
    if r <= 0:
        draw.rectangle([x1, y1, x2, y2], fill=fill)
        return
    draw.pieslice([x1, y1, x1 + 2 * r, y1 + 2 * r], 180, 270, fill=fill)
    draw.pieslice([x2 - 2 * r, y1, x2, y1 + 2 * r], 270, 360, fill=fill)
    draw.pieslice([x2 - 2 * r, y2 - 2 * r, x2, y2], 0, 90, fill=fill)
    draw.pieslice([x1, y2 - 2 * r, x1 + 2 * r, y2], 90, 180, fill=fill)
    draw.rectangle([x1 + r, y1, x2 - r, y2], fill=fill)
    draw.rectangle([x1, y1 + r, x2, y2 - r], fill=fill)


def draw_logo(size: int, ss: int = 8) -> Image.Image:
    """绘制指定尺寸的 Logo（透明背景）。ss 为超采样倍数。"""
    n = size * ss
    img = Image.new("RGBA", (n, n), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # 1) 深色圆角徽章（铺满画布，形成稳定外轮廓）
    _round_rect(draw, (0, 0, n - 1, n - 1), radius=int(n * 0.22), fill=BADGE_COLOR)

    # 2) 白色吊牌（竖向，模拟 4.5×9cm 价格牌比例）
    tw = int(n * 0.42)
    th = int(n * 0.60)
    tx1 = (n - tw) // 2
    ty1 = int(n * 0.23)
    tx2 = tx1 + tw
    ty2 = ty1 + th
    _round_rect(draw, (tx1, ty1, tx2, ty2), radius=int(n * 0.10), fill=CARD_COLOR)

    # 3) 挂孔（透出徽章底色，形成"打孔"效果）
    hr = int(n * 0.055)
    hcx = (tx1 + tx2) // 2
    hcy = ty1 + int(th * 0.17)
    draw.ellipse([hcx - hr, hcy - hr, hcx + hr, hcy + hr], fill=BADGE_COLOR)

    # 4) 两行价格信息（第二行更短，模拟价格数字）
    lw = max(1, int(n * 0.055))
    lx1 = tx1 + int(tw * 0.22)
    lx2 = tx2 - int(tw * 0.22)
    ly1 = ty1 + int(th * 0.47)
    ly2 = ty1 + int(th * 0.64)
    draw.line([(lx1, ly1), (lx2, ly1)], fill=BADGE_COLOR, width=lw)
    draw.line([(lx1, ly2), (lx1 + int((lx2 - lx1) * 0.62), ly2)], fill=BADGE_COLOR, width=lw)

    # 超采样降采样，保证小尺寸平滑锐利
    return img.resize((size, size), Image.LANCZOS)


def main():
    # 高分辨率源图
    draw_logo(256).save(HERE / "logo_256.png", "PNG")
    # 顶栏小图（22px）
    draw_logo(22).save(HERE / "logo_22.png", "PNG")
    # 多分辨率 ICO（16~256），Pillow 会据此内嵌各尺寸
    sizes = [16, 20, 24, 32, 48, 64, 128, 256]
    draw_logo(256).save(
        HERE / "logo.ico", "ICO", sizes=[(s, s) for s in sizes]
    )
    print("已生成：")
    for p in ["logo_256.png", "logo_22.png", "logo.ico"]:
        print("  -", HERE / p)


if __name__ == "__main__":
    main()
