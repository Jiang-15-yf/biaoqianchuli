"""core.label_image - 生成带尺寸标注的标签参考图（嵌入 Excel）

原 HTML 用 Canvas 实现；本模块用 Pillow 等价实现，输出 PNG bytes。
"""
from __future__ import annotations

import io
from typing import Iterable, Optional

from PIL import Image, ImageDraw, ImageFont

# 标注色与原始 JS 版保持一致
ANNOTATION_COLOR = "#dc2626"


def _load_font(size: int) -> ImageFont.ImageFont:
    """尝试加载常见中文字体，失败时回退默认"""
    candidates = [
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/msyhbd.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/simsun.ttc",
        "/System/Library/Fonts/PingFang.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    return ImageFont.load_default()


def _draw_size_annotation(d: ImageDraw.ImageDraw, W: int, H: int, w_cm: float, h_cm: float) -> None:
    """在画布 d 上叠加尺寸标注（边框 + 底部宽文字 + 中心大文字）"""
    pad = 12
    # 外框
    d.rectangle([pad, pad, W - pad, H - pad], outline=ANNOTATION_COLOR, width=2)
    # 底部宽标注（箭头线 + 文字）
    arrow_y = H - pad - 6
    d.line([(pad + 10, arrow_y), (W - pad - 10, arrow_y)], fill=ANNOTATION_COLOR, width=2)
    d.line([(pad + 10, H - pad - 10), (pad + 10, H - pad - 2)], fill=ANNOTATION_COLOR, width=2)
    d.line([(W - pad - 10, H - pad - 10), (W - pad - 10, H - pad - 2)], fill=ANNOTATION_COLOR, width=2)
    font_big = _load_font(18)
    txt = f"{w_cm:.1f} cm"
    bbox = d.textbbox((0, 0), txt, font=font_big)
    d.text(((W - bbox[2]) / 2, H - pad - bbox[3] - 14), txt, fill=ANNOTATION_COLOR, font=font_big)
    # 中心大文字
    font_center = _load_font(22)
    txt2 = f"{w_cm:.1f} × {h_cm:.1f} cm"
    bbox = d.textbbox((0, 0), txt2, font=font_center)
    d.text(((W - bbox[2]) / 2, (H - bbox[3]) / 2), txt2, fill=ANNOTATION_COLOR, font=font_center)


def make_label_diagram(w_cm: float, h_cm: float) -> bytes:
    """生成空白示意图（无参考图时使用），返回 PNG bytes"""
    W = max(200, round(w_cm * 90))
    H = max(200, round(h_cm * 90))
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    _draw_size_annotation(d, W, H, w_cm, h_cm)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def make_annotated_from_bytes(
    image_bytes: bytes, w_cm: float, h_cm: float, annotate: bool = False
) -> bytes:
    """将标签图按目标宽高比裁剪填充并返回 PNG bytes

    注意：用户上传的参考图通常已自带尺寸标注，因此默认不再叠加额外标注，
    仅做等比缩放/填充，保证嵌入 Excel 后不变形、不裁切主要内容。
    annotate=True 时（自定义尺寸开关打开），叠加红色尺寸标注：
    外框 + 底部宽度标注 + 中心「宽 × 高 cm」大字。
    """
    src = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    target_ratio = w_cm / h_cm
    # 计算目标画布尺寸（保持原图比例与目标比例对齐）
    if src.width / src.height > target_ratio:
        H = round(src.width / target_ratio)
        W = src.width
    else:
        W = round(src.height * target_ratio)
        H = src.height
    # 等比缩放并居中，确保原图内容完整不被裁切
    # 背景白色，避免 target_ratio 与原图比例不一致时露黑边
    draw_ratio = min(W / src.width, H / src.height)
    bg = Image.new("RGB", (W, H), "white")
    dw, dh = src.width * draw_ratio, src.height * draw_ratio
    bg.paste(src, (round((W - dw) / 2), round((H - dh) / 2)))
    if annotate:
        _draw_size_annotation(ImageDraw.Draw(bg), W, H, w_cm, h_cm)
    buf = io.BytesIO()
    bg.save(buf, format="PNG")
    return buf.getvalue()


# ----------------------------------------------------------------
# 默认空白标签图：去除烙印的旧尺寸标注，重绘新尺寸
# ----------------------------------------------------------------

# 默认图（301×456）里旧标注的相对位置（比例，方便复用到同构资源）
_OLD_TOP_LINE_Y = 33 / 456      # 顶部横向标注线 y
_OLD_LEFT_LINE_X = 62 / 301     # 左侧纵向标注线 x
_FRAME_TOP_Y = 35 / 456         # 标签框顶
_FRAME_LEFT_X = 76 / 301        # 标签框左
_TEXT_COLOR_FALLBACK = (210, 50, 45)


def _is_red(p) -> bool:
    r, g, b = p[0], p[1], p[2]
    return r > 130 and (r - g) > 60 and (r - b) > 60


def _make_blank_tag_ref(image_bytes: bytes, w_cm: float, h_cm: float) -> bytes:
    """在默认空白标签图上标注「新」尺寸：先擦除旧的红色尺寸标记，再重绘。

    旧标记 = 图中烙印的红色「4,5 cm / 9 cm」文字与三条红色标注线；
    黑色的绳子、标签框、DeFacto 字样全部保留。
    """
    img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    W, H = img.size
    px = img.load()

    # 采样原标注的红色（擦除前）
    reds: list = []
    for y in range(0, min(H, round(H * 0.08)), 2):
        for x in range(0, W, 2):
            if _is_red(px[x, y]):
                reds.append(px[x, y])
    if reds:
        red = max(set(reds), key=reds.count)
    else:
        red = _TEXT_COLOR_FALLBACK

    # 1) 擦除红色像素：顶部条带（文字+横线）、左侧条带（文字+纵线）、右缘线
    top_y = round(H * _FRAME_TOP_Y)
    left_x = round(W * _FRAME_LEFT_X)
    for y in range(H):
        for x in range(W):
            if not _is_red(px[x, y]):
                continue
            in_top = y < top_y
            in_left = x < left_x
            in_right = x >= W - 4
            if in_top or in_left or in_right:
                px[x, y] = (255, 255, 255)

    d = ImageDraw.Draw(img)
    line_y = round(H * _OLD_TOP_LINE_Y)
    line_x = round(W * _OLD_LEFT_LINE_X)
    # 2) 重绘标注线（与原样式一致：顶部横线 + 左/右纵线）
    d.line([(round(W * 0.27), line_y), (round(W * 0.93), line_y)], fill=red, width=2)
    d.line([(line_x, top_y), (line_x, round(H * 0.97))], fill=red, width=2)
    d.line([(W - 2, 0), (W - 2, H)], fill=red, width=2)
    # 3) 新尺寸文字（样式与原文字一致）
    font_txt = _load_font(max(14, round(H * 0.042)))
    txt_w = f"{w_cm:g} cm"
    bbox = d.textbbox((0, 0), txt_w, font=font_txt)
    d.text((round(W * 0.60 - bbox[2] / 2), line_y - bbox[3] - 4), txt_w, fill=red, font=font_txt)
    txt_h = f"{h_cm:g} cm"
    bbox = d.textbbox((0, 0), txt_h, font=font_txt)
    d.text((round(W * 0.105 - bbox[2] / 2), round(H * 0.52 - bbox[3] / 2)), txt_h, fill=red, font=font_txt)

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def blank_tag_reference(image_bytes: Optional[bytes], w_cm: float, h_cm: float, custom_size: bool) -> bytes:
    """合并表「空白标签参考图」的统一入口

    - custom_size=False：图保持原样（默认图自带 4.5×9cm 标注；用户参考图不叠加）
    - custom_size=True：默认图 → 擦旧标注、写新尺寸；用户参考图 → 叠加红色尺寸标注
    """
    if custom_size:
        if image_bytes:
            return make_annotated_from_bytes(image_bytes, w_cm, h_cm, annotate=True)
        default = _load_default_tag_bytes()
        if default:
            return _make_blank_tag_ref(default, w_cm, h_cm)
        return make_label_diagram(w_cm, h_cm)
    # 开关关闭：保持原样
    if image_bytes:
        return make_annotated_from_bytes(image_bytes, w_cm, h_cm, annotate=False)
    default = _load_default_tag_bytes()
    if default:
        return make_annotated_from_bytes(default, w_cm, h_cm, annotate=False)
    return make_label_diagram(w_cm, h_cm)


def _load_default_tag_bytes() -> Optional[bytes]:
    """读取内置默认空白标签图（assets/default_blank_tag.png）"""
    import os
    path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        "assets", "default_blank_tag.png")
    try:
        with open(path, "rb") as f:
            return f.read()
    except Exception:
        return None