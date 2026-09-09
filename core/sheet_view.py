# -*- coding: utf-8 -*-
"""core.sheet_view - 用 Canvas 渲染 SheetModel，做到「所见即所得」

渲染效果对齐 Excel：
- 顶部列标 A/B/C…，左侧行号 1/2/3…
- 单元格 1px 网格线 + 底色填充（S码绿 / 合计黄）
- 合并单元格按范围一次性绘制（视觉上与真实导出一致）
- L10 锚点的标签参考图按模型尺寸绘制

坐标：内容区从行列表头内侧开始；列宽 = 模型字符宽 × 7 + 6，行高 = 模型点高 × 4/3。
"""
from __future__ import annotations

from io import BytesIO
from tkinter import Canvas, ttk
from typing import Dict, List

from PIL import Image, ImageTk

from openpyxl.utils import column_index_from_string, get_column_letter


# 颜色
GRID = "#D9D9DE"
HDR_BG = "#F2F3F5"
HDR_FG = "#6B6F76"
CORNER_BG = "#E9EAED"
TEXT = "#1C1C1E"

ROW_HDR_W = 40
COL_HDR_H = 22


def _col_letter(n: int) -> str:
    return get_column_letter(n)


def _to_rgb(color: str) -> str:
    """openpyxl 用 8 位 ARGB（如 FFFF00），Tk 画布只认 6 位 RGB。取末 6 位并加 #。"""
    if not color:
        return "#FFFFFF"
    c = color.lstrip("#")
    if len(c) == 8:
        c = c[2:]
    return "#" + c


def _fit(text: str, max_px: int, cjk_w: int, ascii_w: int) -> str:
    """按估计字符宽度截断文本，超出加省略号"""
    if not text:
        return ""
    out: List[str] = []
    cur = 0
    for ch in str(text):
        w = cjk_w if ord(ch) > 0x2E80 else ascii_w
        if cur + w > max_px:
            return "".join(out) + "…"
        out.append(ch)
        cur += w
    return "".join(out)


class SheetView(ttk.Frame):
    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        self.model = None
        self._img_refs: List[ImageTk.PhotoImage] = []

        self.canvas = Canvas(self, bg="#FFFFFF", highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.hsb.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.canvas.bind("<MouseWheel>", self._on_wheel)
        self.canvas.bind("<Shift-MouseWheel>", self._on_wheel_shift)

    def _on_wheel(self, e):
        self.canvas.yview_scroll(-1 * (e.delta // 120), "units")

    def _on_wheel_shift(self, e):
        self.canvas.xview_scroll(-1 * (e.delta // 120), "units")

    def set_model(self, model) -> None:
        self.model = model
        self._img_refs.clear()
        self._draw()

    def clear(self) -> None:
        self.model = None
        self._img_refs.clear()
        self.canvas.delete("all")
        self.canvas.configure(scrollregion=(0, 0, 1, 1))

    # ------------------------------------------------------------
    def _geometry(self):
        model = self.model
        col_w: Dict[int, int] = {}
        for c in range(1, model.n_cols + 1):
            col_w[c] = max(40, round(model.col_widths.get(c, 8.43) * 7 + 6))
        row_h: Dict[int, int] = {}
        for r in range(1, model.n_rows + 1):
            row_h[r] = max(18, round(model.row_heights.get(r, 15) * 4 / 3))
        return col_w, row_h

    def _draw(self):
        self.canvas.delete("all")
        if not self.model:
            self.canvas.configure(scrollregion=(0, 0, 1, 1))
            return
        model = self.model
        col_w, row_h = self._geometry()

        col_x: Dict[int, int] = {0: 0}
        for c in range(1, model.n_cols + 1):
            col_x[c] = col_x[c - 1] + col_w[c]
        row_y: Dict[int, int] = {0: 0}
        for r in range(1, model.n_rows + 1):
            row_y[r] = row_y[r - 1] + row_h[r]

        content_w = col_x[model.n_cols]
        content_h = row_y[model.n_rows]
        total_w = ROW_HDR_W + content_w
        total_h = COL_HDR_H + content_h
        self.canvas.configure(scrollregion=(0, 0, total_w + 2, total_h + 2))

        cv = self.canvas
        font_data = ("Microsoft YaHei", 10)
        font_bold = ("Microsoft YaHei", 10, "bold")
        font_hdr = ("Microsoft YaHei", 9, "bold")

        cv.create_rectangle(0, 0, total_w + 2, total_h + 2, fill="#FFFFFF", outline="")

        # 合并覆盖表：被合并区域中（非左上）的单元格需跳过
        covered = set()
        for (r1, c1, r2, c2) in model.merges:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if (r, c) != (r1, c1):
                        covered.add((r, c))

        # ---- 单元格 ----
        for (r, c), cell in model.cells.items():
            if (r, c) in covered:
                continue
            x0 = ROW_HDR_W + col_x[c - 1]
            y0 = COL_HDR_H + row_y[r - 1]
            span = (1, 1)
            for (mr1, mc1, mr2, mc2) in model.merges:
                if mr1 == r and mc1 == c:
                    span = (mr2 - mr1 + 1, mc2 - mc1 + 1)
                    break
            w = col_w[c]
            h = row_h[r]
            if span != (1, 1):
                w = sum(col_w[k] for k in range(c, c + span[1]))
                h = sum(row_h[k] for k in range(r, r + span[0]))
            fill = _to_rgb(cell.fill) if cell.fill else "#FFFFFF"
            cv.create_rectangle(x0, y0, x0 + w, y0 + h, fill=fill, outline=GRID, width=1)
            txt = _fit(cell.value, w - 10, cjk_w=10, ascii_w=6) if cell.value not in (None, "") else ""
            if txt:
                anchor = "w" if cell.align == "left" else ("e" if cell.align == "right" else "center")
                tx = x0 + 5 if anchor == "w" else (x0 + w - 5 if anchor == "e" else x0 + w / 2)
                cv.create_text(tx, y0 + h / 2, text=txt, anchor=anchor,
                              font=font_bold if cell.bold else font_data, fill=TEXT)

        # ---- 列标头 ----
        for c in range(1, model.n_cols + 1):
            x0 = ROW_HDR_W + col_x[c - 1]
            cv.create_rectangle(x0, 0, x0 + col_w[c], COL_HDR_H, fill=HDR_BG, outline=GRID, width=1)
            cv.create_text(x0 + col_w[c] / 2, COL_HDR_H / 2, text=_col_letter(c),
                          font=font_hdr, fill=HDR_FG)

        # ---- 行标头 ----
        for r in range(1, model.n_rows + 1):
            y0 = COL_HDR_H + row_y[r - 1]
            cv.create_rectangle(0, y0, ROW_HDR_W, y0 + row_h[r], fill=HDR_BG, outline=GRID, width=1)
            cv.create_text(ROW_HDR_W / 2, y0 + row_h[r] / 2, text=str(r), font=font_hdr, fill=HDR_FG)

        # ---- 左上角 ----
        cv.create_rectangle(0, 0, ROW_HDR_W, COL_HDR_H, fill=CORNER_BG, outline=GRID, width=1)

        # ---- 参考图 ----
        for anchor, img_bytes, iw, ih in model.images:
            try:
                col_letter = "".join(ch for ch in anchor if ch.isalpha())
                row_no = int("".join(ch for ch in anchor if ch.isdigit()))
                cidx = column_index_from_string(col_letter)
                x0 = ROW_HDR_W + col_x.get(cidx - 1, 0)
                y0 = COL_HDR_H + row_y.get(row_no - 1, 0)
                im = Image.open(BytesIO(img_bytes)).convert("RGBA")
                im = im.resize((int(iw), int(ih)), Image.LANCZOS)
                tkim = ImageTk.PhotoImage(im)
                self._img_refs.append(tkim)
                cv.create_image(x0, y0, anchor="nw", image=tkim)
            except Exception as e:
                print(f"[warn] 预览绘制图片失败: {e}")
