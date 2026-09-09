"""core.pdf_rearrange - PDF 价格签自动检测与重排

核心算法（与原 HTML/JS 版本对齐）：
1. render_page_pix  : PyMuPDF 把页面渲染成高分辨率 RGB 位图（DETECT_SCALE=3）
2. detect_stickers  : 行/列投影 + 阈值 RGB<235 → 候选矩形 → merge_boxes 合并 → 加 PAD
3. merge_boxes      : 先按 x 中心聚列，再按 y 聚行，最后按 y 全局二次合并
4. rearrange_pdf    : 把检测到的标签按 PDF 源页面区域直接绘制到新页面（保清晰度）

居中与尺寸统一策略（v1.3 起）：
- 全部标签的标称尺寸取「中位数 + 容差吸收抖动」，得到一个统一的 crop_w × crop_h；
- 每枚标签的裁切框 = 以它自己的墨迹中心 (cx, cy) 为圆心、尺寸恒为 crop_w × crop_h；
- 目标槽位尺寸与裁切框相同，缩放恒为 1.0，直接居中摆放。
结果：所有输出标签尺寸 100% 一致，且每枚标签的墨迹中心严格落在输出槽位中心。
裁切框越出源页边界时（标签贴边），只绘制页面内的部分，越界处留白，居中关系不变。

单位约定：
- 检测坐标：DETECT_SCALE=3 下像素（仅检测阶段使用）
- 输出坐标：PDF 点（pt，72pt=1inch），与 PyMuPDF 一致
- 转换：pt = px / DETECT_SCALE

v1.6 新增：
- 替换每枚标签左上角的 FSC 标记。
  资产 assets/fsc_block.png：由参考 P260901923.pdf 第 1 页左上角 FSC 块
  （树标+FSC+MIX+Paper+FSC™）以 16x 高清渲染并擦除原编号区域得到。
  编号由调用方传入 fsc_number（如 "C202518"），用 Helvetica-Bold 矢量
  重绘，与参考文件位置/字号一致。
"""
from __future__ import annotations

import math
import os
import re
import statistics
from collections import Counter
from dataclasses import dataclass, field
from typing import List, Tuple, Dict, Any, Optional

import fitz  # PyMuPDF
from PIL import Image

from .units import cm_to_pt, mm_to_pt, pt_to_cm

# 检测固定放大倍率（与原 JS 一致，DETECT_SCALE 像素→pt 需除以此值）
DETECT_SCALE = 3
# 检测框外扩像素（≈ 10pt ≈ 3.5mm，避免切边）
PAD_PX = 30
# 墨色阈值（原 JS：RGB 任意通道 < 235 即视为有墨）
INK_THRESHOLD = 235
# 合并阈值（DETECT_SCALE 像素）
COL_GAP = 90
ROW_GAP = 180
# 行/列投影最小宽度阈值
MIN_BOX_W = 40
MIN_BAND_H = 40
GAP_TOLERANCE = 15  # 投影间断容忍行数
# 小条带并入大条带的最大间隙（DETECT_SCALE 像素，60px ≈ 20pt ≈ 7mm）。
# 用于吸收紧贴标签边缘的小字（如标签底部的打印日期行），
# 同时排除离标签较远的行间说明文字（如 Printing Quantity）。
SMALL_BAND_GAP = 60

# 尺码顺序（重排时款式内按此顺序排列）
SIZE_ORDER = ["XS", "S", "M", "L", "XL", "XXL", "3XL", "XXXXL"]
_SIZE_SET = set(SIZE_ORDER)
# 默认防切边（mm）：太小容易切到标签边缘内容，不建议 < 3
DEFAULT_PADDING_MM = 3.0
# 输出页面（标签物理尺寸）默认值，参考样例 P260901923.pdf（页面 ≈ 4.54×7.54cm）
DEFAULT_LABEL_W_CM = 4.5
DEFAULT_LABEL_H_CM = 7.5
# 页面自动下限：内容墨迹 + 该余量（mm），保证任何情况下标签内容不被页面裁掉
PAGE_SAFETY_MM = 2.0
# 颜色识别关键词（命中行中最后一个 "/" 后的词即颜色名，如 "Ürün çeşidi: .../Ekru"）
COLOR_KEYWORDS = ("çeşidi", "cesidi", "çeşit", "renk", "identity", "colour", "color")

# ================================================================
# FSC 标记（v1.6）：每页标签左上角替换
# ================================================================
# 资产 assets/fsc_block.png：以参考 P260901923.pdf 第 1 页左侧 FSC 块的完整
# 墨迹区域（树标+™ + FSC + MIX + Paper + FSC™，16x 高清渲染）并擦除原编号
# "C202518" 得到。块在参考页面中的区域：x 11.0→32.5、y 22.8→50.8。
# 参考文件中 FSC 图像放置矩形为 (11.28, 24.35, 32.26, 49.02)，故块原点相对
# 图像矩形左上角偏移 DX=-0.28、DY=-1.55（树顶与 ™ 高出图像矩形）。
FSC_BLOCK_W_PT = 21.5
FSC_BLOCK_H_PT = 28.0
FSC_BLOCK_DX = -0.28   # 块原点 = FSC 图像矩形左上角 + (DX, DY)
FSC_BLOCK_DY = -1.55
# 矢量编号相对块左上角的偏移（参考文件中 "C202518" 实测：起点 x=21.08、
# 基线 y=50.0，块原点 (11.0, 22.8)）
FSC_TEXT_X_OFF = 10.08
FSC_TEXT_BASELINE_Y_OFF = 27.2
# 数字 cap 高 ≈ 1.9pt，对应 Helvetica（常规体）字号 ≈ 2.6pt。
# 改用 Arial Narrow Bold 后笔画更细窄，可适当放大字号2.8。黑体simhei选用2.2
FSC_TEXT_FONTSIZE_PT = 2.8
# 编号最长宽度上限：超过则按比例缩小字号
FSC_TEXT_MAX_WIDTH_PT = FSC_BLOCK_W_PT - FSC_TEXT_X_OFF - 0.4
# 默认编号（参考值，与 P260901923.pdf 原图一致）
DEFAULT_FSC_NUMBER = "C202518"
# Arial Narrow Bold 路径（Windows 系统字体；项目 assets 下有副本做兜底）
_FSC_FONT_CANDIDATES = (
    r"C:\Windows\Fonts\ARIALNB.TTF",     #Arial Narrow Bold字体
    #r"C:\Windows\Fonts\simhei.ttf",        #黑体
    #r"C:\Windows\Fonts\msyhbd.ttc",        #微软雅黑粗体
    os.path.join(os.path.dirname(__file__), "..", "assets", "ARIALNB.ttf"),
)

def _fsc_font_path() -> str:
    for p in _FSC_FONT_CANDIDATES:
        if os.path.exists(p):
            return p
    return _FSC_FONT_CANDIDATES[0]  # 找不到时返回首选路径，让 PyMuPDF 自己报错

# 源 PDF 中识别 FSC 图像的尺寸范围（pt）。FSC logo 为 112×132 px @ 96 dpi ≈
# 20.98×24.67pt；判断允许 ±25% 容差，避免不同 dpi 嵌入产生的轻微差异。
_FSC_SRC_IMG_W = 20.98
_FSC_IMG_W_RANGE = (_FSC_SRC_IMG_W * 0.75, _FSC_SRC_IMG_W * 1.25)
_FSC_IMG_H_RANGE = (24.0, 30.0)
_FSC_IMG_ASPECT_RANGE = (0.65, 0.95)  # 宽/高 ≈ 0.85


def _load_fsc_block_png() -> Optional[bytes]:
    """加载 assets/fsc_block.png（已擦除原编号的高清 FSC 块位图）。"""
    import sys
    from pathlib import Path
    candidates = []
    if hasattr(sys, "_MEIPASS"):
        candidates.append(Path(sys._MEIPASS) / "assets" / "fsc_block.png")
    candidates.append(Path(__file__).resolve().parent.parent / "assets" / "fsc_block.png")
    for p in candidates:
        try:
            with open(p, "rb") as f:
                return f.read()
        except OSError:
            continue
    return None


def _find_fsc_rect_on_page(page: "fitz.Page", ink_rect_pt: fitz.Rect) -> Optional[fitz.Rect]:
    """在源页面里找出位于 ink_rect 内的 FSC 图像矩形。

    ink_rect_pt：标签墨迹矩形，与 collect_labels 一致使用 top-down 坐标系
    （y 向下，原点在页面左上角）。get_image_rects 同样返回 top-down 坐标，
    因此可以直接做相交判断。

    返回：FSC 图像矩形（top-down），或 None。
    """
    for img in page.get_images(full=True):
        try:
            rects = page.get_image_rects(img[0])
        except Exception:
            continue
        for r in rects:
            w, hgt = r.width, r.height
            if not (_FSC_IMG_W_RANGE[0] <= w <= _FSC_IMG_W_RANGE[1]):
                continue
            if not (_FSC_IMG_H_RANGE[0] <= hgt <= _FSC_IMG_H_RANGE[1]):
                continue
            aspect = w / hgt if hgt else 0
            if not (_FSC_IMG_ASPECT_RANGE[0] <= aspect <= _FSC_IMG_ASPECT_RANGE[1]):
                continue
            cx, cy = r.x0 + w / 2, r.y0 + hgt / 2
            if not (ink_rect_pt.x0 <= cx <= ink_rect_pt.x1 and ink_rect_pt.y0 <= cy <= ink_rect_pt.y1):
                continue
            return r
    return None


def _stamp_fsc(
    out_page: "fitz.Page",
    src_page: "fitz.Page",
    src_crop_pt: fitz.Rect,
    target_pt: fitz.Rect,
    fsc_block_png: bytes,
    fsc_number: str,
) -> None:
    """在输出页标签左上角盖 FSC 块（覆盖原块、重绘图片与编号）。

    src_crop_pt ：该标签在源页中的裁切框（top-down，y 向下）。
    target_pt    ：输出页槽位矩形（y-up，PDF pt；与 src_crop_pt 坐标系不同）。
    fsc_block_png：擦除原编号后的 FSC 位图（21.5×28.0pt）。
    fsc_number   ：如 "C202518"，空白时跳过。

    坐标说明（重要）：
    - src_crop_pt、src_fsc 都是 top-down；target_pt 是 y-up。
    - 缩放恒为 1.0，因此把源坐标的位移 (src_fsc - src_crop) 直接加到 target
      上即可，方向差异对纯平移没影响（与 _draw_label 的 sub 矩形同款约定）。
    - 输出矩形在 y-up 中：y0 = 下沿，y1 = 上沿；
      块 "在 top-down 中向下延伸 FSC_BLOCK_H_PT" 对应 y-up 中 "向上延伸同值"。
    """
    if not fsc_number:
        return
    src_fsc = _find_fsc_rect_on_page(src_page, src_crop_pt)
    if src_fsc is not None:
        # 块原点 = FSC 图像矩形左上角 + 固定偏移（含树顶/™ 超出图像矩形的部分）
        src_block_x0 = src_fsc.x0 + FSC_BLOCK_DX
        src_block_y0 = src_fsc.y0 + FSC_BLOCK_DY
    else:
        # 回落：以墨迹（top-down）左上角为图像矩形近似位置，再加同款偏移
        src_block_x0 = src_crop_pt.x0 + FSC_BLOCK_DX
        src_block_y0 = src_crop_pt.y0 + FSC_BLOCK_DY
    # 平移量 = 源坐标位移（数值上同目标，因为缩放 1.0）
    dx = src_block_x0 - src_crop_pt.x0
    dy = src_block_y0 - src_crop_pt.y0
    # 输出矩形（y-up：y0=下沿）
    bx0 = target_pt.x0 + dx
    by0 = target_pt.y0 + dy
    bx1 = bx0 + FSC_BLOCK_W_PT
    by1 = by0 + FSC_BLOCK_H_PT
    # 与输出页面求交（标签若贴边，越界部分自然留空）
    page_rect_yup = fitz.Rect(0, 0, out_page.rect.width, out_page.rect.height)
    block = fitz.Rect(bx0, by0, bx1, by1) & page_rect_yup
    if block.width <= 0 or block.height <= 0:
        return
    # 1. 用白底覆盖原区域（含原图与原矢量编号）
    out_page.draw_rect(block, color=None, fill=(1, 1, 1), overlay=True)
    # 2. 贴新 FSC 位图（21.5×28.0pt，与 assets/fsc_block.png 同比例）
    img_rect = fitz.Rect(bx0, by0, bx0 + FSC_BLOCK_W_PT, by0 + FSC_BLOCK_H_PT)
    out_page.insert_image(img_rect, stream=fsc_block_png, overlay=True)
    # 3. 重绘编号（Arial Narrow Bold 矢量字，与参考文件字重一致）
    fontfile = _fsc_font_path()
    avail_w = FSC_TEXT_MAX_WIDTH_PT
    fontsize = FSC_TEXT_FONTSIZE_PT
    try:
        txt = fsc_number
        # 缩放计算：用 fitz.Font 准确测量（get_text_length 不支持 fontfile）
        fobj = fitz.Font(fontfile=fontfile)
        for _ in range(8):
            tw = fobj.text_length(txt, fontsize=fontsize)
            if tw <= avail_w or fontsize <= 0.5:
                break
            fontsize *= avail_w / tw
        tx = bx0 + FSC_TEXT_X_OFF
        ty = by0 + FSC_TEXT_BASELINE_Y_OFF
        out_page.insert_text(
            (tx, ty), txt,
            fontfile=fontfile, fontsize=fontsize,
            color=(0, 0, 0), overlay=True,
        )
    except Exception:
        pass


def size_index(size: Optional[str]) -> int:
    """尺码 → 排序序号；无法识别的排在最后"""
    if not size:
        return len(SIZE_ORDER)
    s = str(size).strip().upper()
    return SIZE_ORDER.index(s) if s in _SIZE_SET else len(SIZE_ORDER)


def detect_label_size(page: "fitz.Page", rect: "fitz.Rect") -> Optional[str]:
    """识别单枚标签的尺码（XS / S / M / L / XL / XXL ...）

    思路：
    1. 取标签区域内的文本词，按行（y 中心±3pt）分组；
    2. 一行里出现 ≥3 个尺码词的，视为「尺码表/多语言表」，跳过；
    3. 其余行中取唯一命中的尺码词，最后取出现次数最多者（多语言标签会重复出现）。
    """
    try:
        words = page.get_text("words", clip=rect)
    except Exception:
        return None
    lines: Dict[int, List[Tuple[float, str]]] = {}
    for w in words:
        x0, y0, x1, y1, txt = w[0], w[1], w[2], w[3], w[4]
        key = round((y0 + y1) / 2 / 3.0)
        lines.setdefault(key, []).append((x0, txt))
    hits: List[str] = []
    for key in sorted(lines):
        toks = [re.sub(r"[^A-Za-z0-9]", "", t).upper() for _, t in sorted(lines[key])]
        found = [t for t in toks if t in _SIZE_SET]
        if len(found) >= 3:
            continue  # 尺码表行，跳过
        if len(found) == 1:
            hits.append(found[0])
    if not hits:
        return None
    cnt = Counter(hits)
    # 出现次数最多者优先；次数相同则取尺码顺序靠前、位置靠前者
    best = max(cnt.items(), key=lambda kv: (kv[1], -size_index(kv[0])))
    return best[0]


def _group_lines(words) -> Dict[int, List[Tuple[float, str]]]:
    """把 (x0, y0, x1, y1, text) 词序列按 y 中心（±3pt）分组成行"""
    lines: Dict[int, List[Tuple[float, str]]] = {}
    for w in words:
        x0, y0, x1, y1, txt = w[0], w[1], w[2], w[3], w[4]
        key = round((y0 + y1) / 2 / 3.0)
        lines.setdefault(key, []).append((x0, txt))
    return lines


def detect_label_color(page: "fitz.Page", rect: "fitz.Rect") -> Optional[str]:
    """识别单枚标签的颜色名

    规则：按行扫描，第一行同时包含颜色关键词与 "/" 的，
    取最后一个 "/" 之后的文字作为颜色名（如 "Ürün çeşidi: .../Ekru" → Ekru）。
    """
    try:
        words = page.get_text("words", clip=rect)
    except Exception:
        return None
    lines = _group_lines(words)
    for key in sorted(lines):
        line = " ".join(t for _, t in sorted(lines[key]))
        low = line.lower()
        if any(k in low for k in COLOR_KEYWORDS) and "/" in line:
            seg = line.rsplit("/", 1)[1].strip().strip(" .:,;|")
            if seg:
                return seg
    return None


@dataclass
class Box:
    """DETECT_SCALE 像素坐标系下的矩形"""
    x: float
    y: float
    w: float
    h: float

    @property
    def cx(self) -> float:
        return self.x + self.w / 2

    @property
    def cy(self) -> float:
        return self.y + self.h / 2

    @property
    def x2(self) -> float:
        return self.x + self.w

    @property
    def y2(self) -> float:
        return self.y + self.h

    def to_dict(self) -> Dict[str, float]:
        return {"x": self.x, "y": self.y, "w": self.w, "h": self.h}


@dataclass
class Detection:
    """一次检测结果：raw 为未加 PAD 的核心框，padded 为加 PAD 后的外扩框"""
    raw: Box
    padded: Box


@dataclass
class Label:
    """单枚标签的源区域信息

    PyMuPDF / Pillow 坐标系：原点(0,0)在页面左上角，y 轴向下增长。

    居中策略：
        裁切框不是「检测框 + 外扩」，而是以标签自身的墨迹中心 (cx, cy) 为圆心、
        以统一标称尺寸 (crop_w × crop_h) 画一个矩形。这样所有标签：
        1) 尺寸完全相同（裁切框大小一致，缩放恒为 1.0）
        2) 内容完全居中（墨迹中心 == 裁切框中心 == 输出槽位中心）
    """
    src_idx: int   # 源 PDF 页索引
    cx: float      # 墨迹中心 x（源页面坐标系 pt）
    cy: float      # 墨迹中心 y（源页面坐标系 pt）
    core_w: float  # 该标签检测出的标称宽（pt，含 PAD_PX 留白）
    core_h: float  # 该标签检测出的标称高（pt，含 PAD_PX 留白）
    # --- 以下用于多文件合并排序（可选） ---
    doc: Any = None        # 源 fitz.Document（批量合并时区分来源文件）
    size: str = ""         # 识别出的尺码（XS / S / M / L / XL / XXL）
    color: str = ""        # 识别出的颜色名（如 Ekru / Lacivert）
    style: str = ""        # 款式名（默认取文件名主干）
    order: int = 0         # 文件内原始顺序，用于稳定排序

    def crop_rect(self, crop_w: float, crop_h: float) -> fitz.Rect:
        """以墨迹中心为圆心生成统一尺寸的裁切框（fitz.Rect: x0,y0 左上 / x1,y1 右下）"""
        return fitz.Rect(
            self.cx - crop_w / 2,
            self.cy - crop_h / 2,
            self.cx + crop_w / 2,
            self.cy + crop_h / 2,
        )


# ================================================================
# 渲染与检测
# ================================================================

def render_page_pix(page: fitz.Page, scale: float = DETECT_SCALE) -> Image.Image:
    """渲染 PDF 页面为 PIL Image（用于检测）"""
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def _ink(img: Image.Image, x: int, y: int) -> bool:
    r, g, b = img.getpixel((x, y))
    return r < INK_THRESHOLD or g < INK_THRESHOLD or b < INK_THRESHOLD


def detect_stickers(page: fitz.Page) -> List[Detection]:
    """对单页 PDF 做投影法检测，返回候选标签列表

    每个 Detection 包含：
    - raw: 未加 PAD 的核心检测框（用于计算输出页面尺寸）
    - padded: 加 PAD 后的外扩框（用于后续裁切，防止切边）
    """
    img = render_page_pix(page)
    W, H = img.size

    # 行投影
    row_cnt = [0] * H
    for y in range(H):
        c = 0
        for x in range(0, W, 2):
            if _ink(img, x, y):
                c += 1
        row_cnt[y] = c
    row_on = lambda y: row_cnt[y] > 2

    # 抽取有墨的水平条带（先不做高度过滤，小条带稍后就近合并）
    bands: List[Tuple[int, int]] = []
    s = -1
    gap = 0
    for y in range(H):
        if row_on(y):
            if s < 0:
                s = y
            gap = 0
        elif s >= 0:
            gap += 1
            if gap > GAP_TOLERANCE:
                if y - gap - s > 0:
                    bands.append((s, y - gap))
                s = -1
    if s >= 0:
        bands.append((s, H - 1))

    # 小条带（孤立小字）就近并入大条带：
    # 紧贴标签边缘的小字（标签底部日期等）属于标签内容，必须并入，
    # 否则示例图/重排会切掉底部时间行；离标签较远的说明文字保持排除。
    big = [[b[0], b[1]] for b in bands if b[1] - b[0] > MIN_BAND_H]
    small = [b for b in bands if b[1] - b[0] <= MIN_BAND_H]
    for sb in small:
        best: Optional[List[int]] = None
        best_d: Optional[int] = None
        for bb in big:
            if sb[1] < bb[0]:
                d = bb[0] - sb[1]
            elif sb[0] > bb[1]:
                d = sb[0] - bb[1]
            else:
                d = 0
            if best_d is None or d < best_d:
                best, best_d = bb, d
        if best is not None and best_d <= SMALL_BAND_GAP:
            best[0] = min(best[0], sb[0])
            best[1] = max(best[1], sb[1])
    bands = [(b[0], b[1]) for b in big]

    # 每个条带内做列投影
    boxes: List[Box] = []
    for y0, y1 in bands:
        col_cnt = [0] * W
        for x in range(W):
            for y in range(y0, y1, 2):
                if _ink(img, x, y):
                    col_cnt[x] += 1
        cs = -1
        for x in range(W + 1):
            on = x < W and col_cnt[x] > 1
            if on and cs < 0:
                cs = x
            if (not on or x == W) and cs >= 0:
                if x - cs > MIN_BOX_W:
                    boxes.append(Box(cs, y0, x - cs, y1 - y0))
                cs = -1

    merged = merge_boxes(boxes)

    # 生成 raw + padded 两种框
    result: List[Detection] = []
    for b in merged:
        nx = max(0, b.x - PAD_PX)
        ny = max(0, b.y - PAD_PX)
        nx2 = min(W, b.x2 + PAD_PX)
        ny2 = min(H, b.y2 + PAD_PX)
        raw = Box(b.x, b.y, b.w, b.h)
        padded = Box(nx, ny, nx2 - nx, ny2 - ny)
        result.append(Detection(raw=raw, padded=padded))
    return result


# ================================================================
# 框合并
# ================================================================

def _merge_group(group: List[Box]) -> Box:
    x = min(b.x for b in group)
    y = min(b.y for b in group)
    x2 = max(b.x2 for b in group)
    y2 = max(b.y2 for b in group)
    return Box(x, y, x2 - x, y2 - y)


def merge_boxes(boxes: List[Box], col_gap: float = COL_GAP, row_gap: float = ROW_GAP) -> List[Box]:
    """合并候选框：
    1) 按 x 中心聚列（容差 col_gap）
    2) 每列内按 y 聚行（容差 row_gap）
    3) 全局按 y 再合并一次，处理列聚类切开的边界
    """
    if not boxes:
        return []
    # 1. 按 x 中心聚列
    sorted_x = sorted(boxes, key=lambda b: b.cx)
    cols: List[Dict[str, Any]] = []
    for b in sorted_x:
        cx = b.cx
        target = None
        for c in cols:
            if abs(cx - c["sum"] / c["n"]) < col_gap:
                target = c
                break
        if target:
            target["boxes"].append(b)
            target["sum"] += cx
            target["n"] += 1
        else:
            cols.append({"boxes": [b], "sum": cx, "n": 1})

    # 2. 列内按 y 合并
    merged: List[Box] = []
    for col in cols:
        by_y = sorted(col["boxes"], key=lambda b: b.y)
        g = [by_y[0]]
        for i in range(1, len(by_y)):
            last = g[-1]
            dy = by_y[i].y - last.y2
            if dy < row_gap:
                g.append(by_y[i])
            else:
                merged.append(_merge_group(g))
                g = [by_y[i]]
        merged.append(_merge_group(g))

    # 3. 全局按 y 二次合并（处理同一标签被列聚类切开的边界）
    by_y = sorted(merged, key=lambda b: (b.y, b.x))
    if not by_y:
        return []
    final: List[Box] = []
    fg = [by_y[0]]
    for i in range(1, len(by_y)):
        last = fg[-1]
        dy = by_y[i].y - last.y2
        dx = abs(by_y[i].cx - last.cx)
        if dy < min(row_gap, 80) and dx < col_gap * 2:
            fg.append(by_y[i])
        else:
            final.append(_merge_group(fg))
            fg = [by_y[i]]
    final.append(_merge_group(fg))
    return final


def _canonical_size(values: List[float], tolerance: float = 1.15) -> float:
    """从一组检测尺寸中求「统一标称尺寸」

    思路：
    - 先取中位数，得到最可能的真实标签尺寸；
    - 再吸收与中位数接近的正常抖动（≤ tolerance 倍），取其中的最大值，
      保证没有任何一枚标签的内容被裁掉；
    - 明显偏大的框（如误把两枚标签合并成一枚）视为检测异常，不参与放大，
      否则会拖着所有标签一起变大。
    """
    if not values:
        return 0.0
    med = statistics.median(values)
    cands = [v for v in values if v <= med * tolerance]
    return max(cands) if cands else med


def _draw_label(
    out_page: "fitz.Page",
    src: "fitz.Document",
    lb: Label,
    crop: fitz.Rect,
    target: fitz.Rect,
) -> None:
    """把源页 crop 区域绘制到输出页 target 区域。

    裁切框可能越出源页边界（标签贴近页边时必然发生）。这里把裁切框与源页求交，
    只绘制落在页面内的部分，并按相同的位移量映射到目标区域的对应子矩形——
    越界部分保持白色，墨迹中心依然严格居中。
    """
    src_doc = lb.doc if lb.doc is not None else src
    page_rect = src_doc[lb.src_idx].rect
    clip = crop & page_rect
    if clip.width <= 0 or clip.height <= 0:
        return
    sub = fitz.Rect(
        target.x0 + (clip.x0 - crop.x0),
        target.y0 + (clip.y0 - crop.y0),
        target.x0 + (clip.x1 - crop.x0),
        target.y0 + (clip.y1 - crop.y0),
    )
    # pymupdf 1.24+ 改为静态方法：fitz.Page.show_pdf_page(page, rect, docsrc, pno, clip=...)
    fitz.Page.show_pdf_page(out_page, sub, src_doc, lb.src_idx, clip=clip)


# ================================================================
# 主入口：PDF 重排
# ================================================================

# ================================================================
# 标签图提取（供 Excel 合并表嵌入「价格牌示例图」）
# ================================================================

def extract_label_png(input_path: str, label_index: int = 0, zoom: int = 3) -> Optional[bytes]:
    """从 PDF 中提取第 label_index 枚标签，渲染为 PNG bytes。

    用于合并 Excel：每个款式取其对应 PDF 的第一枚价格牌作为示例图。
    检测失败 / 无标签时返回 None（调用方自行跳过）。
    """
    try:
        src = fitz.open(input_path)
    except Exception:
        return None
    try:
        count = 0
        for p_idx in range(len(src)):
            dets = detect_stickers(src[p_idx])
            for det in dets:
                if count == label_index:
                    r = det.raw
                    # 与重排裁切一致的 PAD 外扩，再与页面求交防越界
                    clip = fitz.Rect(
                        (r.x - PAD_PX) / DETECT_SCALE,
                        (r.y - PAD_PX) / DETECT_SCALE,
                        (r.x2 + PAD_PX) / DETECT_SCALE,
                        (r.y2 + PAD_PX) / DETECT_SCALE,
                    ) & src[p_idx].rect
                    pix = src[p_idx].get_pixmap(
                        matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=False)
                    return pix.tobytes("png")
                count += 1
    except Exception:
        return None
    finally:
        src.close()
    return None


def collect_labels(doc: "fitz.Document", style: str = "") -> List[Label]:
    """检测单个 PDF 中的全部标签，识别尺码与颜色后按「颜色 → 尺码顺序」稳定排序

    排序规则：
        1. 颜色按标签在 PDF 中首次出现的顺序（即原 PDF 颜色页顺序）；
        2. 同一颜色内按尺码顺序：XS → S → M → L → XL → XXL；
        3. 颜色、尺码均无法识别的标签按原顺序排在对应位置。
    """
    labels: List[Label] = []
    for p_idx in range(len(doc)):
        dets = detect_stickers(doc[p_idx])
        for det in dets:
            r = det.raw
            rect = fitz.Rect(
                r.x / DETECT_SCALE, r.y / DETECT_SCALE,
                r.x2 / DETECT_SCALE, r.y2 / DETECT_SCALE,
            )
            size = detect_label_size(doc[p_idx], rect)
            color = detect_label_color(doc[p_idx], rect)
            labels.append(Label(
                src_idx=p_idx,
                cx=(r.x + r.w / 2) / DETECT_SCALE,
                cy=(r.y + r.h / 2) / DETECT_SCALE,
                core_w=det.padded.w / DETECT_SCALE,
                core_h=det.padded.h / DETECT_SCALE,
                doc=doc,
                size=size or "",
                color=color or "",
                style=style,
                order=len(labels),
            ))
    # 颜色分组序号：按首次出现顺序
    color_rank: Dict[str, int] = {}

    def ckey(lb: Label) -> str:
        return lb.color if lb.color else f"__页{lb.src_idx + 1}"

    for lb in labels:
        ck = ckey(lb)
        if ck not in color_rank:
            color_rank[ck] = len(color_rank)
    labels.sort(key=lambda l: (color_rank[ckey(l)], size_index(l.size), l.order))
    return labels


def rearrange_pdfs(
    input_paths: List[str],
    output_path: str,
    per_page: int = 1,
    margin_mm: float = 0,
    padding_mm: float = DEFAULT_PADDING_MM,
    label_w_cm: float = DEFAULT_LABEL_W_CM,
    label_h_cm: float = DEFAULT_LABEL_H_CM,
    fsc_number: str = "",
) -> Dict[str, Any]:
    """主入口：多个 PDF 检测 + 合并重排为**一个** PDF

    排序规则：
        1. 按输入文件顺序（每个文件 = 一个款式）依次往后排；
        2. 款式内按颜色分组，颜色按首次出现顺序；
        3. 同一颜色内按尺码顺序：XS → S → M → L → XL → XXL。

    页面布局（参考样例 P260901923.pdf：每页一枚标签、页面 ≈ 标签物理尺寸）：
        - 输出页面 = 标签物理尺寸（默认 4.5×7.5cm），页边距默认 0；
        - 页面会自动放大到 ≥「内容墨迹 + 2mm」，保证内容绝不被页面裁掉；
        - 防切边 padding_mm 只扩大裁切范围（出血），不改变页面尺寸：
          裁切区域以墨迹中心为准 1:1 居中绘制，超出页面的部分自然裁掉。

    参数：
        input_paths : 源 PDF 路径列表（按款式顺序传入）
        output_path : 输出 PDF 路径（单个汇总文件）
        per_page    : 每页输出几枚（默认 1）
        margin_mm   : 输出页面四周留白（毫米）
        padding_mm  : 防切边余量（毫米），默认 3，不建议 < 3
        label_w_cm  : 标签（页面）宽 cm，默认 4.5
        label_h_cm  : 标签（页面）高 cm，默认 7.5
        fsc_number  : v1.6 - 每页左上角 FSC 编号（如 "C202518"）；空串则不加盖

    返回：{'labels', 'pages', 'page_w_cm', 'page_h_cm', 'label_w_cm', 'label_h_cm',
           'files': [{'path','style','labels','color_groups':[{'color','sizes'}]}],
           'fsc_number': 实际加盖的编号（便于空时区分）}
    """
    docs: List["fitz.Document"] = []
    groups: List[Tuple[str, str, List[Label]]] = []
    try:
        for path in input_paths:
            doc = fitz.open(path)
            docs.append(doc)
            style = os.path.splitext(os.path.basename(path))[0]
            groups.append((path, style, collect_labels(doc, style)))
    except Exception:
        for d in docs:
            d.close()
        raise

    labels: List[Label] = [l for _, _, ls in groups for l in ls]
    if not labels:
        for d in docs:
            d.close()
        raise ValueError("未检测到标签（请检查 PDF 内容或调整合并阈值）")

    # 内容墨迹尺寸（padded 框 - PAD 留白）；各文件分别求标称值（抗抖动）后取最大
    pad_pt = PAD_PX / DETECT_SCALE
    ink_w = max(_canonical_size([l.core_w - pad_pt * 2 for l in ls]) for _, _, ls in groups if ls)
    ink_h = max(_canonical_size([l.core_h - pad_pt * 2 for l in ls]) for _, _, ls in groups if ls)

    # 页面 = 标签物理尺寸，自动下限 = 墨迹 + PAGE_SAFETY_MM
    label_w = max(cm_to_pt(label_w_cm), ink_w + mm_to_pt(PAGE_SAFETY_MM))
    label_h = max(cm_to_pt(label_h_cm), ink_h + mm_to_pt(PAGE_SAFETY_MM))

    # 裁切（出血）= 墨迹 + 2×防切边；绘制 1:1，超出页面的部分自然裁掉
    padding_pt = mm_to_pt(padding_mm)
    crop_w = ink_w + padding_pt * 2
    crop_h = ink_h + padding_pt * 2

    margin_pt = mm_to_pt(margin_mm)
    slot_gap = 6.0  # 多枚/页时槽位间固定间隔（pt）

    page_w = label_w + margin_pt * 2
    if per_page == 1:
        page_h = label_h + margin_pt * 2
    else:
        page_h = per_page * label_h + (per_page - 1) * slot_gap + margin_pt * 2

    n_pages = math.ceil(len(labels) / per_page)

    # v1.6：FSC 资源按需加载；空串/缺资产则不加盖
    fsc_png = _load_fsc_block_png() if fsc_number else None
    if fsc_number and not fsc_png:
        print("[warn] 未找到 assets/fsc_block.png，跳过 FSC 加盖")

    out = fitz.open()
    try:
        for _ in range(n_pages):
            out.new_page(width=page_w, height=page_h)
        # 注意：PyMuPDF 中缓存的 Page 引用会丢失 parent，必须用 out[i] 重新取

        for i, lb in enumerate(labels):
            # 裁切框：以该标签墨迹中心为圆心，尺寸恒为 crop_w × crop_h
            crop = lb.crop_rect(crop_w, crop_h)
            # 目标槽位：标签区域中心（裁切框同尺寸 1:1 居中摆放，出血越出页面即裁掉）
            slot = i % per_page
            cx = margin_pt + label_w / 2
            cy = margin_pt + slot * (label_h + slot_gap) + label_h / 2
            target = fitz.Rect(cx - crop_w / 2, cy - crop_h / 2, cx + crop_w / 2, cy + crop_h / 2)
            out_page = out[i // per_page]
            _draw_label(out_page, lb.doc, lb, crop, target)
            # v1.6：左上角加盖 FSC（覆盖原块、贴图、重绘编号）
            if fsc_png:
                src_page = (lb.doc if lb.doc is not None else docs[0])[lb.src_idx]
                _stamp_fsc(out_page, src_page, crop, target, fsc_png, fsc_number)

        out.save(output_path)
    finally:
        out.close()
        for d in docs:
            d.close()

    return {
        "labels": len(labels),
        "pages": n_pages,
        "page_w_cm": pt_to_cm(page_w),
        "page_h_cm": pt_to_cm(page_h),
        "label_w_cm": pt_to_cm(label_w),
        "label_h_cm": pt_to_cm(label_h),
        "fsc_number": fsc_number,
        "files": [
            {
                "path": p,
                "style": st,
                "labels": len(ls),
                "color_groups": _color_groups(ls),
            }
            for p, st, ls in groups
        ],
    }


def _color_groups(labels: List[Label]) -> List[Dict[str, Any]]:
    """把排好序的标签按颜色聚组，输出 [{'color': 名, 'sizes': [尺码...]}]"""
    out: List[Dict[str, Any]] = []
    for lb in labels:
        ck = lb.color if lb.color else "未识别"
        if out and out[-1]["color"] == ck:
            out[-1]["sizes"].append(lb.size or "-")
        else:
            out.append({"color": ck, "sizes": [lb.size or "-"]})
    return out


def rearrange_pdf(
    input_path: str,
    output_path: str,
    per_page: int = 1,
    margin_mm: float = 0,
    padding_mm: float = DEFAULT_PADDING_MM,
    label_w_cm: float = DEFAULT_LABEL_W_CM,
    label_h_cm: float = DEFAULT_LABEL_H_CM,
    fsc_number: str = "",
) -> Dict[str, Any]:
    """单文件重排（内部走批量入口，同样按颜色/尺码排序）"""
    return rearrange_pdfs(
        [input_path], output_path,
        per_page=per_page, margin_mm=margin_mm, padding_mm=padding_mm,
        label_w_cm=label_w_cm, label_h_cm=label_h_cm,
        fsc_number=fsc_number,
    )