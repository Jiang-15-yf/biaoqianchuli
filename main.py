# -*- coding: utf-8 -*-
"""main - Tkinter 桌面 GUI 入口（极简风格 v1.3）

布局：顶栏 + 左侧导航 + 内容区 + 底栏。
- Excel 标签：选择表 → 参数 → 预览（电子表格画布，与导出完全一致）→ 导出
- PDF 重排：选择 PDF → 解析 → 参数 → 生成

业务逻辑全部在 core/* 中，本文件只做 UI 编排。
"""
from __future__ import annotations

import os
import sys
import traceback
from pathlib import Path
from typing import Optional
import tkinter as tk
from tkinter import font as tkfont
from tkinter import Tk, StringVar, DoubleVar, IntVar, BooleanVar, END, DISABLED, NORMAL, Button, Frame, Label, Checkbutton
from tkinter import Entry as TkEntry
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from PIL import Image, ImageTk

from core.pdf_rearrange import (rearrange_pdf, rearrange_pdfs, detect_stickers,
                                DETECT_SCALE, DEFAULT_PADDING_MM)
from core.excel_pricecard import (
    parse_table, build_pricecard, export_xlsx, build_sheet_model,
    build_merged_model, export_merged_xlsx,
)
from core.sheet_view import SheetView

# 拖拽支持（tkinterdnd2 已 vendor 到项目 vendor/ 目录，缺失时自动降级为仅按钮选择）
_VENDOR_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vendor")
if os.path.isdir(_VENDOR_DIR) and _VENDOR_DIR not in sys.path:
    sys.path.insert(0, _VENDOR_DIR)


APP_TITLE = "价格牌处理工具"
APP_VERSION = "1.6"

# ================================================================
# 资源路径（兼容开发目录与 PyInstaller 单文件打包）
# ================================================================
def _resource_path(rel: str) -> str:
    """返回资源文件绝对路径。PyInstaller 单文件运行时资源在 _MEIPASS。"""
    if hasattr(sys, "_MEIPASS"):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).parent.resolve()
    return str(base / rel)


# ================================================================
# 配色（极简 · 近黑 + 发丝级灰线 + 大量留白）
# ================================================================
INK = "#17181C"          # 主文字 / 主按钮
INK_2 = "#5C616B"        # 次要文字
INK_3 = "#9AA0A8"        # 弱文字
LINE = "#ECECF0"         # 发丝边框
LINE_2 = "#F5F5F7"       # 浅填充
BG = "#FFFFFF"           # 内容背景
BG_APP = "#F6F6F8"       # 应用背景
SIDEBAR = "#FAFAFB"      # 侧栏背景
ACCENT = "#2F6BFF"       # 仅用于激活态指示的克制强调
SELECT = "#EFF1F4"       # 选中/悬停浅底

FONT_TITLE = ("Microsoft YaHei", 14, "bold")
FONT_SUB = ("Microsoft YaHei", 10)
FONT_NAV = ("Microsoft YaHei", 11)
FONT_LABEL = ("Microsoft YaHei", 10)
FONT_BOLD = ("Microsoft YaHei", 10, "bold")
FONT_MONO = ("Consolas", 10)
FONT_HINT = ("Microsoft YaHei", 9)
FONT_BUTTON = ("Microsoft YaHei", 10)

CORNER = 8  # 圆角半径（px）


# ================================================================
# 圆角控件
# ================================================================

class RoundedButton(tk.Canvas):
    """圆角按钮：primary=近黑填充；secondary=白底发丝边框。

    用 Canvas 绘制圆角矩形 + 居中文字，规避 tk.Button 无法圆角的问题。
    """

    def __init__(self, parent, text="", kind="primary", command=None, state="normal",
                 padx=18, pady=8, font=FONT_BUTTON, bg=BG, **kw):
        super().__init__(parent, bd=0, highlightthickness=0, bg=bg, **kw)
        self._text = text
        self._kind = kind
        self._command = command
        self._state = state
        self._padx = padx
        self._pady = pady
        self._font = font
        self._hover = False
        self._items = []
        f = tkfont.Font(family=font[0], size=font[1],
                        weight=font[2] if len(font) > 2 else "normal")
        w = f.measure(text) + padx * 2
        h = f.metrics("linespace") + pady * 2
        self._bw, self._bh = max(int(w), 44), int(h)
        self.config(width=self._bw, height=self._bh)
        self._draw()
        self.bind("<Enter>", lambda e: self._enter())
        self.bind("<Leave>", lambda e: self._leave())
        self.bind("<Button-1>", lambda e: self._press())
        self.bind("<ButtonRelease-1>", lambda e: self._release())
        self.config(cursor="hand2" if state == "normal" else "arrow")

    def _colors(self):
        if self._state == DISABLED:
            return {"fill": "#EDEEF0", "outline": "#EDEEF0", "fg": "#C7C9CE"}
        if self._kind == "primary":
            fill = "#000000" if self._hover else INK
            return {"fill": fill, "outline": fill, "fg": "#FFFFFF"}
        fill = SELECT if self._hover else BG
        return {"fill": fill, "outline": LINE, "fg": INK}

    def _draw(self):
        for it in self._items:
            self.delete(it)
        self._items = []
        c = self._colors()
        self._items.extend(self._round_rect(1, 1, self._bw - 2, self._bh - 2, CORNER,
                                            fill=c["fill"], outline=c["outline"], width=1))
        self._items.append(self.create_text(self._bw / 2, self._bh / 2, text=self._text,
                                            fill=c["fg"], font=self._font, anchor="center"))

    def _round_rect(self, x1, y1, x2, y2, r, fill=BG, outline=LINE, width=1):
        """绘制圆角矩形：先填充（外描边与填充同色，避免内部出现横线），再描边。"""
        its = []
        kw_fill = dict(fill=fill, outline=fill, width=1)
        its.append(self.create_arc(x1, y1, x1 + 2 * r, y1 + 2 * r, start=90, extent=90, style="pieslice", **kw_fill))
        its.append(self.create_arc(x2 - 2 * r, y1, x2, y1 + 2 * r, start=0, extent=90, style="pieslice", **kw_fill))
        its.append(self.create_arc(x2 - 2 * r, y2 - 2 * r, x2, y2, start=270, extent=90, style="pieslice", **kw_fill))
        its.append(self.create_arc(x1, y2 - 2 * r, x1 + 2 * r, y2, start=180, extent=90, style="pieslice", **kw_fill))
        its.append(self.create_rectangle(x1 + r, y1, x2 - r, y2, **kw_fill))
        its.append(self.create_rectangle(x1, y1 + r, x2, y2 - r, **kw_fill))
        if outline and width > 0:
            kw_arc = dict(fill="", outline=outline, width=width)
            kw_line = dict(fill=outline, width=width)
            its.append(self.create_arc(x1, y1, x1 + 2 * r, y1 + 2 * r, start=90, extent=90, style="arc", **kw_arc))
            its.append(self.create_arc(x2 - 2 * r, y1, x2, y1 + 2 * r, start=0, extent=90, style="arc", **kw_arc))
            its.append(self.create_arc(x2 - 2 * r, y2 - 2 * r, x2, y2, start=270, extent=90, style="arc", **kw_arc))
            its.append(self.create_arc(x1, y2 - 2 * r, x1 + 2 * r, y2, start=180, extent=90, style="arc", **kw_arc))
            its.append(self.create_line(x1 + r, y1, x2 - r, y1, **kw_line))
            its.append(self.create_line(x1 + r, y2, x2 - r, y2, **kw_line))
            its.append(self.create_line(x1, y1 + r, x1, y2 - r, **kw_line))
            its.append(self.create_line(x2, y1 + r, x2, y2 - r, **kw_line))
        return its

    def _enter(self):
        if self._state == DISABLED:
            return
        self._hover = True
        self._draw()

    def _leave(self):
        self._hover = False
        self._draw()

    def _press(self):
        if self._state == DISABLED or not self._command:
            return

    def _release(self):
        if self._state == DISABLED or not self._command:
            return
        self._command()

    def set_state(self, state):
        self._state = state
        self.config(cursor="hand2" if state == "normal" else "arrow")
        self._draw()


class RoundedField(tk.Frame):
    """圆角输入框：Canvas 绘制圆角边框，内部 Entry 无边框嵌入。"""

    def __init__(self, parent, var, width=9, font=FONT_LABEL, bg=BG, **kw):
        super().__init__(parent, bg=bg, **kw)
        self._bg = bg
        self._canvas = tk.Canvas(self, bg=bg, bd=0, highlightthickness=0)
        self._canvas.place(relx=0, rely=0, relwidth=1, relheight=1)
        self._entry = TkEntry(self, textvariable=var, width=width, font=font,
                              bg=bg, fg=INK, relief="flat", bd=0,
                              insertbackground=INK, highlightthickness=0)
        self._entry.pack(padx=10, pady=5)
        self._focus = False
        self._entry.bind("<FocusIn>", lambda e: self._redraw(True))
        self._entry.bind("<FocusOut>", lambda e: self._redraw(False))
        self._canvas.bind("<Configure>", lambda e: self._redraw(self._focus))

    def _redraw(self, focus):
        self._focus = focus
        self._canvas.delete("all")
        w = self._canvas.winfo_width()
        h = self._canvas.winfo_height()
        if w < 4 or h < 4:
            return
        color = ACCENT if focus else LINE
        r = CORNER
        self._round_rect(1, 1, w - 2, h - 2, r, fill=self._bg,
                         outline=color, width=1 if not focus else 1.5)

    def _round_rect(self, x1, y1, x2, y2, r, fill=BG, outline=LINE, width=1):
        c = self._canvas
        kw_fill = dict(fill=fill, outline=fill, width=1)
        c.create_arc(x1, y1, x1 + 2 * r, y1 + 2 * r, start=90, extent=90, style="pieslice", **kw_fill)
        c.create_arc(x2 - 2 * r, y1, x2, y1 + 2 * r, start=0, extent=90, style="pieslice", **kw_fill)
        c.create_arc(x2 - 2 * r, y2 - 2 * r, x2, y2, start=270, extent=90, style="pieslice", **kw_fill)
        c.create_arc(x1, y2 - 2 * r, x1 + 2 * r, y2, start=180, extent=90, style="pieslice", **kw_fill)
        c.create_rectangle(x1 + r, y1, x2 - r, y2, **kw_fill)
        c.create_rectangle(x1, y1 + r, x2, y2 - r, **kw_fill)
        if outline and width > 0:
            kw_arc = dict(fill="", outline=outline, width=width)
            kw_line = dict(fill=outline, width=width)
            c.create_arc(x1, y1, x1 + 2 * r, y1 + 2 * r, start=90, extent=90, style="arc", **kw_arc)
            c.create_arc(x2 - 2 * r, y1, x2, y1 + 2 * r, start=0, extent=90, style="arc", **kw_arc)
            c.create_arc(x2 - 2 * r, y2 - 2 * r, x2, y2, start=270, extent=90, style="arc", **kw_arc)
            c.create_arc(x1, y2 - 2 * r, x1 + 2 * r, y2, start=180, extent=90, style="arc", **kw_arc)
            c.create_line(x1 + r, y1, x2 - r, y1, **kw_line)
            c.create_line(x1 + r, y2, x2 - r, y2, **kw_line)
            c.create_line(x1, y1 + r, x1, y2 - r, **kw_line)
            c.create_line(x2, y1 + r, x2, y2 - r, **kw_line)


# ================================================================
# 通用工具
# ================================================================

def _set_status(text_widget, msg: str) -> None:
    text_widget.config(state=NORMAL)
    text_widget.delete("1.0", END)
    if msg:
        text_widget.insert("1.0", msg)
    text_widget.config(state=DISABLED)


def _btn(parent, text: str, kind: str = "primary", command=None, state: str = "normal", width=None) -> RoundedButton:
    """圆角按钮：primary=近黑填充；secondary=白底发丝边框。"""
    return RoundedButton(parent, text=text, kind=kind, command=command, state=state)


def set_btn_state(btn, state: str) -> None:
    if hasattr(btn, "set_state"):
        btn.set_state(state)
    else:
        btn.config(state=state)


def _field(parent, label: str, var, hint: str = "", width: int = 9) -> Frame:
    """一行参数：标签 + 圆角输入框 + 提示"""
    row = Frame(parent, bg=BG)
    Label(row, text=label, font=FONT_LABEL, bg=BG, fg=INK_2).pack(side="left")
    field = RoundedField(row, var, width=width)
    field.pack(side="left", padx=8)
    if hint:
        Label(row, text=hint, font=FONT_HINT, bg=BG, fg=INK_3).pack(side="left")
    return row


# ================================================================
# 主窗口
# ================================================================

class App:
    def __init__(self, root: Tk) -> None:
        self.root = root
        root.title(f"{APP_TITLE}  v{APP_VERSION}")
        root.geometry("1120x860")
        root.minsize(940, 720)
        root.configure(bg=BG_APP)
        try:
            ttk.Style().theme_use("clam")
        except Exception:
            pass

        self._build_topbar()
        self._build_sidebar()
        self._build_content_host()
        self._build_statusbar()

        self.current = None
        self._nav_to("excel")
        self._setup_dragdrop()

    # ----------------------------------------------------------
    # 拖入导入
    # ----------------------------------------------------------
    def _setup_dragdrop(self) -> None:
        """把文件/图片直接拖进窗口即完成选择（需要 tkinterdnd2，缺失时静默跳过）"""
        try:
            from tkinterdnd2 import DND_FILES
        except Exception:
            self.dnd_ok = False
            return
        try:
            self.dnd_ok = True
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            self.dnd_ok = False

    def _on_drop(self, event) -> None:
        try:
            paths = [str(p) for p in self.root.tk.splitlist(event.data)]
        except Exception:
            return
        if not paths:
            return
        imgs = [p for p in paths if os.path.splitext(p)[1].lower()
                in (".png", ".jpg", ".jpeg", ".bmp")]
        xls = [p for p in paths if os.path.splitext(p)[1].lower()
               in (".xls", ".xlsx", ".csv")]
        pdfs = [p for p in paths if os.path.splitext(p)[1].lower() == ".pdf"]
        other = len(paths) - len(imgs) - len(xls) - len(pdfs)
        if self.current == "pdf":
            if pdfs:
                self._pd_set_files(pdfs)
            if xls:
                messagebox.showinfo("提示", "PDF 重排页签只接受 PDF 文件；Excel 请切换到「Excel 合并转化」页签")
        else:
            if xls:
                self._ex_set_files(xls)
            if pdfs:
                messagebox.showinfo("提示", "Excel 合并转化不再使用价格签 PDF；PDF 请切换到「PDF 标签重排」页签")
            if imgs:
                self._ex_set_image(imgs[0])
        if other:
            self.status_var.set("部分文件类型不受支持，已忽略")

    # ----------------------------------------------------------
    # 框架
    # ----------------------------------------------------------
    def _build_topbar(self) -> None:
        bar = Frame(self.root, bg=BG, height=60)
        bar.pack(fill="x", side="top")
        bar.pack_propagate(False)
        inner = Frame(bar, bg=BG)
        inner.pack(fill="x", expand=True, padx=24)
        # 标志：Logo（价格签形状）
        try:
            self.logo_img = ImageTk.PhotoImage(
                Image.open(_resource_path("assets/logo_22.png")).convert("RGBA")
            )
            logo_lbl = Label(inner, image=self.logo_img, bg=BG)
            logo_lbl.pack(side="left", padx=(0, 12))
        except Exception as e:
            # 兜底：如果 Logo 资源缺失，仍用近黑方块保证界面不崩
            mark = Frame(inner, bg=INK, width=22, height=22)
            mark.pack(side="left", padx=(0, 12))
            print(f"[warn] Logo 加载失败: {e}")
        # 标题（pack 自动计算宽度，避免 place 导致的截断）
        Label(inner, text=APP_TITLE, font=FONT_TITLE, bg=BG, fg=INK).pack(side="left")
        # 版本靠右
        Label(inner, text=f"v{APP_VERSION}", font=FONT_HINT, bg=BG, fg=INK_3).pack(side="right")
        # 底线
        sep = Frame(self.root, bg=LINE, height=1)
        sep.pack(fill="x", side="top")

    def _build_sidebar(self) -> None:
        side = Frame(self.root, bg=SIDEBAR, width=245)
        side.pack(fill="y", side="left")
        side.pack_propagate(False)
        Label(side, text="功能", font=FONT_HINT, bg=SIDEBAR, fg=INK_3).pack(anchor="w", padx=20, pady=(22, 10))

        self.nav_items = {}
        self._nav_row(side, "excel", "①  Excel 合并转化（多文件）")
        self._nav_row(side, "pdf", "②  PDF 价格签重排（批量）")

    def _nav_row(self, side, key, text) -> None:
        item = Frame(side, bg=SIDEBAR)
        item.pack(fill="x", pady=2)
        bar = Frame(item, bg=SIDEBAR, width=3)
        bar.pack(side="left", fill="y", padx=(16, 0))
        btn = Button(item, text=text, font=FONT_NAV, bg=SIDEBAR, fg=INK_2,
                     bd=0, relief="flat", anchor="w", padx=10, pady=8,
                     activebackground=SIDEBAR, activeforeground=INK,
                     command=lambda k=key: self._nav_to(k))
        btn.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.nav_items[key] = (item, bar, btn)

    def _build_content_host(self) -> None:
        self.host = Frame(self.root, bg=BG)
        self.host.pack(fill="both", expand=True, side="left")

        self.excel_tab = Frame(self.host, bg=BG)
        self.pdf_tab = Frame(self.host, bg=BG)
        self._build_excel_tab(self.excel_tab)
        self._build_pdf_tab(self.pdf_tab)

    def _build_statusbar(self) -> None:
        bar = Frame(self.root, bg=BG, height=30)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        sep = Frame(self.root, bg=LINE, height=1)
        sep.pack(fill="x", side="bottom")
        self.status_var = StringVar(value="就绪")
        Label(bar, textvariable=self.status_var, font=FONT_HINT, bg=BG, fg=INK_2).pack(side="left", padx=24)

    def _nav_to(self, key: str) -> None:
        for k, (item, bar, btn) in self.nav_items.items():
            active = (k == key)
            btn.config(bg=BG if active else SIDEBAR, fg=INK if active else INK_2)
            bar.config(bg=INK if active else SIDEBAR)
        if self.current and self.current != key:
            getattr(self, f"{self.current}_tab").pack_forget()
        self.current = key
        tab = getattr(self, f"{key}_tab")
        tab.pack(fill="both", expand=True, padx=0, pady=0)
        self.status_var.set("就绪")

    # ----------------------------------------------------------
    # 区块工具
    # ----------------------------------------------------------
    def _section(self, parent, title: str, subtitle: str = "") -> Frame:
        """内容区的一个区块：标题 + 分隔线 + 内容 Frame"""
        wrap = Frame(parent, bg=BG)
        wrap.pack(fill="x", padx=28, pady=(8, 4))
        head = Frame(wrap, bg=BG)
        head.pack(fill="x")
        Label(head, text=title, font=("Microsoft YaHei", 12, "bold"), bg=BG, fg=INK).pack(side="left")
        if subtitle:
            Label(head, text=subtitle, font=FONT_HINT, bg=BG, fg=INK_3).pack(side="left", padx=10)
        Line = Frame(wrap, bg=LINE, height=1)
        Line.pack(fill="x", pady=(6, 10))
        return wrap

    # ----------------------------------------------------------
    # Excel 标签 Tab
    # ----------------------------------------------------------
    def _build_excel_tab(self, parent: Frame) -> None:
        # 输入：多个 Excel + 可选参考图（v1.6 不再需要价格签PDF）
        sec = self._section(parent, "输入", "可多选 Excel / CSV / XLS · 支持拖入窗口 · 合并为一个 xlsx")
        frow = Frame(sec, bg=BG)
        frow.pack(fill="x", pady=(0, 4))
        _btn(frow, "选择Excel文件", "secondary", self._ex_pick_files).pack(side="left")
        # 文件 chip 列表（每个文件一枚 chip，点 × 可单独删除）
        self.ex_chips = Frame(sec, bg=BG)
        self.ex_chips.pack(fill="x", pady=(0, 8))
        self.ex_chips_hint = Label(self.ex_chips, text="未选择文件（可多选）", font=FONT_LABEL, bg=BG, fg=INK_3, anchor="w")
        self.ex_chips_hint.pack(side="left")

        irow = Frame(sec, bg=BG)
        irow.pack(fill="x")
        _btn(irow, "选择参考图", "secondary", self._ex_pick_image).pack(side="left")
        # 参考图 chip 列表（v1.6 单图也支持 × 移除，与网页版样式一致）
        self.ex_img_chips = Frame(sec, bg=BG)
        self.ex_img_chips.pack(fill="x", pady=(4, 0))
        self.ex_img_chips_hint = Label(self.ex_img_chips,
                                        text="未选择（使用默认空白标签图 4.5×9cm）",
                                        font=FONT_LABEL, bg=BG, fg=INK_3, anchor="w")
        self.ex_img_chips_hint.pack(side="left")

        # 参数
        sec2 = self._section(parent, "参数")
        self.ex_rate = DoubleVar(value=4.0)
        self.tag_w = DoubleVar(value=4.5)
        self.tag_h = DoubleVar(value=9.0)
        self.ex_custom_size = BooleanVar(value=False)
        prow1 = Frame(sec2, bg=BG)
        prow1.pack(fill="x", pady=(0, 10))
        _field(prow1, "生产数上浮 %", self.ex_rate, "(0~20)").pack(side="left")
        Checkbutton(prow1, text="自定义标签尺寸", variable=self.ex_custom_size,
                    command=self._ex_toggle_size, font=FONT_LABEL, bg=BG, fg=INK_2,
                    activebackground=BG, highlightthickness=0).pack(side="left", padx=(28, 0))
        prow2 = Frame(sec2, bg=BG)
        prow2.pack(fill="x")
        _field(prow2, "标签宽 cm", self.tag_w, "(推荐 4.5)").pack(side="left", padx=(0, 28))
        _field(prow2, "标签高 cm", self.tag_h, "(推荐 9.0)").pack(side="left")
        # 开关默认关闭：尺寸输入置灰（默认 4.5 / 9.0，不额外标注）
        self._size_rows = (prow2, )
        self._ex_toggle_size()

        # 操作（右对齐：预览 / 导出）
        arow = Frame(parent, bg=BG)
        arow.pack(fill="x", padx=28, pady=12)
        self.ex_go_btn = _btn(arow, "导出合并 .xlsx", "primary", self._ex_generate, state=DISABLED)
        self.ex_go_btn.pack(side="right")
        _btn(arow, "预览", "secondary", self._ex_preview).pack(side="right", padx=(0, 10))

        # 预览区
        sec4 = self._section(parent, "工作表预览", "与导出的合并 .xlsx 完全一致")
        prev_wrap = Frame(parent, bg=BG)
        prev_wrap.pack(fill="both", expand=True, padx=28, pady=(0, 18))
        border = Frame(prev_wrap, bg=LINE, bd=1, relief="solid")
        border.pack(fill="both", expand=True)
        self.ex_sheet = SheetView(border)
        self.ex_sheet.pack(fill="both", expand=True, padx=1, pady=1)

        self.ex_files: list = []          # 多个 Excel 路径
        self.ex_ref_image_bytes = None
        self.ex_img_path: Optional[str] = None  # 当前参考图完整路径（None=用默认）

    def _render_chips(self, files, chips, hint, on_change, on_move=None) -> None:
        """通用文件 chip 列表：圆角胶囊 + × 按钮。"""
        for w in chips.winfo_children():
            if w is not hint:  # 删除全部文件，不销毁 hint
                w.destroy()
        if not files:
            hint.configure(text="未选择文件（可多选）", fg=INK_3, bg=BG)
            hint.pack(side="left")
            return
        hint.pack_forget()
        for idx, p in enumerate(files):
            chip = Frame(chips, bg=LINE_2, bd=0, highlightthickness=1,
                         highlightbackground=LINE, highlightcolor=LINE)
            chip.pack(side="top", anchor="w", padx=0, pady=2)  # ← 顺带改竖排，见下方说明
            inner_pad = Frame(chip, bg=LINE_2, padx=8, pady=3)
            inner_pad.pack()

            # ↓↓↓ 新增：上移 / 下移（首尾自动禁用）
            if on_move is not None:
                for arrow, delta, enable in (
                        ("↑", -1, idx > 0),
                        ("↓", 1, idx < len(files) - 1),
                ):
                    b = Button(inner_pad, text=arrow, font=("Segoe UI", 9),
                               bg=LINE_2, fg=INK_2 if enable else LINE,
                               relief="flat", bd=0, width=2, height=1,
                               cursor="hand2" if enable else "arrow",
                               state=NORMAL if enable else DISABLED,
                               command=lambda i=idx, d=delta: on_move(i, d))
                    b.pack(side="left", padx=(0, 2))

            xbtn = Button(inner_pad, text="×", font=("Segoe UI", 11, "bold"),
                          bg=LINE_2, fg=INK_2, relief="flat", bd=0,
                          activebackground=LINE, activeforeground=INK,
                          width=2, height=1, cursor="hand2",
                          command=lambda pp=p: on_change([f for f in files if f != pp]))
            xbtn.pack(side="left", padx=(0, 4))
            lbl = Label(inner_pad, text=os.path.basename(p), font=FONT_LABEL,
                        bg=LINE_2, fg=INK)
            lbl.pack(side="left")
            for w in (chip, inner_pad, lbl):
                w.bind("<Button-1>", lambda e, pp=p: on_change([f for f in files if f != pp]))
                w.bind("<Enter>", lambda e, c=chip, i=inner_pad, lb=lbl, x=xbtn: (
                    c.config(bg="#E5E7EB"), i.config(bg="#E5E7EB"),
                    lb.config(bg="#E5E7EB"), x.config(bg="#E5E7EB")))
                w.bind("<Leave>", lambda e, c=chip, i=inner_pad, lb=lbl, x=xbtn: (
                    c.config(bg=LINE_2), i.config(bg=LINE_2),
                    lb.config(bg=LINE_2), x.config(bg=LINE_2)))

    # ---- Excel 页签：多文件选择与合并 ----

    def _ex_render_chips(self) -> None:
        self._render_chips(self.ex_files, self.ex_chips, self.ex_chips_hint, self._ex_remove_files)

    def _ex_remove_files(self, remain: list) -> None:
        self.ex_files = remain
        if self.ex_files:
            set_btn_state(self.ex_go_btn, NORMAL)
            self.status_var.set(f"剩余 {len(self.ex_files)} 个文件")
        else:
            set_btn_state(self.ex_go_btn, DISABLED)
            self.status_var.set("已移除全部 Excel 文件")
        self._ex_render_chips()

    def _ex_toggle_size(self) -> None:
        """标签尺寸开关：开 → 尺寸输入可用并在空白标签图上标注新尺寸；
        关 → 默认 4.5 / 9.0，输入置灰、不做任何标注。"""
        on = self.ex_custom_size.get()
        state = NORMAL if on else DISABLED
        for row in getattr(self, "_size_rows", ()):
            for w in row.winfo_children():
                for e in w.winfo_children():
                    if isinstance(e, TkEntry):
                        e.config(state=state)
   
    def _ex_pick_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择 Excel/CSV 文件（可多选，也可直接拖入窗口）",
            filetypes=[("Excel/CSV", "*.xls *.xlsx *.csv"), ("All", "*.*")],
        )
        if paths:
            self._ex_set_files(list(paths))

    # def _ex_set_files(self, paths: list) -> None:
    #     self.ex_files = list(paths)
    #     ok, fails = [], []
    #     for p in self.ex_files:
    #         try:
    #             parse_table(p)
    #             ok.append(p)
    #         except Exception as e:
    #             fails.append(f"{os.path.basename(p)}：{e}")
    #     self.ex_files = ok
    #     if ok:
    #         set_btn_state(self.ex_go_btn, NORMAL)
    #         self.status_var.set(f"已读取 {len(ok)} 个文件" + (f"；{len(fails)} 个失败" if fails else ""))
    #     else:
    #         set_btn_state(self.ex_go_btn, DISABLED)
    #         self.status_var.set("读取失败")
    #     self._ex_render_chips()
    #     if fails:
    #         messagebox.showerror("部分文件读取失败", "\n".join(fails))

    def _ex_set_files(self, paths: list, append: bool = True) -> None:
        cur = list(self.ex_files) if append else []
        for p in paths:
            if p not in cur:  # 去重
                cur.append(p)
        self.ex_files = cur

        ok, fails = [], []
        for p in self.ex_files:
            try:
                parse_table(p)
                ok.append(p)
            except Exception as e:
                fails.append(f"{os.path.basename(p)}：{e}")
        self.ex_files = ok
        if ok:
            set_btn_state(self.ex_go_btn, NORMAL)
            self.status_var.set(f"已累计 {len(ok)} 个文件" + (f"；{len(fails)} 个失败" if fails else ""))
        else:
            set_btn_state(self.ex_go_btn, DISABLED)
            self.status_var.set("读取失败")
        self._ex_render_chips()
        if fails:
            messagebox.showerror("部分文件读取失败", "\n".join(fails))

    def _ex_render_img_chip(self) -> None:
        for w in self.ex_img_chips.winfo_children():
            w.destroy()
        if not self.ex_img_path:
            self.ex_img_chips_hint.configure(text="未选择（使用默认空白标签图 4.5×9cm）",
                                             fg=INK_3, bg=BG)
            self.ex_img_chips_hint.pack(side="left")
            return
        self.ex_img_chips_hint.pack_forget()
        path = self.ex_img_path
        chip = Frame(self.ex_img_chips, bg=LINE_2, bd=0, highlightthickness=1,
                     highlightbackground=LINE, highlightcolor=LINE)
        chip.pack(side="left", pady=2)
        inner = Frame(chip, bg=LINE_2, padx=8, pady=3)
        inner.pack()
        xbtn = Button(inner, text="×", font=("Segoe UI", 11, "bold"),
                      bg=LINE_2, fg=INK_2, relief="flat", bd=0,
                      activebackground=LINE, activeforeground=INK,
                      width=2, height=1, cursor="hand2",
                      command=self._ex_clear_image)
        xbtn.pack(side="left", padx=(0, 4))
        lbl = Label(inner, text=os.path.basename(path), font=FONT_LABEL,
                    bg=LINE_2, fg=INK)
        lbl.pack(side="left")
        for w in (chip, inner, lbl):
            w.bind("<Button-1>", lambda e: self._ex_clear_image())
            w.bind("<Enter>", lambda e, c=chip, i=inner, lb=lbl, x=xbtn: (
                c.config(bg="#E5E7EB"), i.config(bg="#E5E7EB"),
                lb.config(bg="#E5E7EB"), x.config(bg="#E5E7EB")))
            w.bind("<Leave>", lambda e, c=chip, i=inner, lb=lbl, x=xbtn: (
                c.config(bg=LINE_2), i.config(bg=LINE_2),
                lb.config(bg=LINE_2), x.config(bg=LINE_2)))

    def _ex_clear_image(self) -> None:
        self.ex_img_path = None
        self.ex_ref_image_bytes = None
        self._ex_render_img_chip()
        self.status_var.set("已移除参考图（恢复使用默认空白标签图）")

    def _ex_pick_image(self) -> None:
        path = filedialog.askopenfilename(
            title="选择标签参考图（也可直接拖入窗口）",
            filetypes=[("Image", "*.png *.jpg *.jpeg *.bmp"), ("All", "*.*")],
        )
        if path:
            self._ex_set_image(path)

    def _ex_set_image(self, path: str) -> None:
        try:
            with open(path, "rb") as f:
                self.ex_ref_image_bytes = f.read()
            self.ex_img_path = path
            self._ex_render_img_chip()
            self.status_var.set(f"已载入参考图：{os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("载入失败", str(e))

    def _ex_build_sections(self) -> list:
        """解析并构建每个文件的竖排数据，返回 sections 列表"""
        import re
        sections = []
        for p in self.ex_files:
            parsed = parse_table(p)
            built = build_pricecard(parsed, self.ex_rate.get())
            styles = sorted({r["style"] for r in parsed["rows"] if r["style"]})
            pos = set()
            for r in parsed["rows"]:
                for po in re.split(r"[,，、]", r.get("po") or ""):
                    po = po.strip()
                    if po:
                        pos.add(po)
            sections.append({
                "style": styles[0] if len(styles) == 1 else (parsed["rows"][0]["style"] if parsed["rows"] else ""),
                "styles": styles,
                "pos": pos,
                "parsed": parsed,
                "built": built,
            })
        return sections

    def _ex_preview(self) -> None:
        if not self.ex_files:
            messagebox.showwarning("提示", "请先选择 Excel 文件")
            return
        try:
            sections = self._ex_build_sections()
            model = build_merged_model(
                sections,
                tag_w_cm=self.tag_w.get(), tag_h_cm=self.tag_h.get(),
                ref_image_bytes=self.ex_ref_image_bytes,
                custom_size=self.ex_custom_size.get(),
            )
            self.ex_sheet.set_model(model)
            total_body = sum(len(s["built"]["body"]) for s in sections)
            total_blank = sum(len(s["built"]["blank_out"]) for s in sections)
            self.status_var.set(
                f"预览：{len(sections)} 个款式 · 主体 {total_body} 行 + 空白 {total_blank} 行 + 总计 1 行"
            )
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("预览失败", str(e))

    def _ex_generate(self) -> None:
        if not self.ex_files:
            return
        default_name = "特殊价格牌-合并.xlsx"
        try:
            sections = self._ex_build_sections()
            if sections and sections[0]["style"]:
                default_name = f"{sections[0]['style']}-特殊价格牌-合并.xlsx"
        except Exception:
            sections = None
        out_path = filedialog.asksaveasfilename(
            title="保存合并 Excel 为",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=default_name,
        )
        if not out_path:
            return
        try:
            if sections is None:
                sections = self._ex_build_sections()
            export_merged_xlsx(
                sections, out_path,
                tag_w_cm=self.tag_w.get(), tag_h_cm=self.tag_h.get(),
                ref_image_bytes=self.ex_ref_image_bytes,
                custom_size=self.ex_custom_size.get(),
            )
            model = build_merged_model(
                sections,
                tag_w_cm=self.tag_w.get(), tag_h_cm=self.tag_h.get(),
                ref_image_bytes=self.ex_ref_image_bytes,
                custom_size=self.ex_custom_size.get(),
            )
            self.ex_sheet.set_model(model)
            self.status_var.set(f"已导出合并 Excel：{out_path}")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("生成失败", str(e))

    # ----------------------------------------------------------
    # PDF 标签 Tab
    # ----------------------------------------------------------
    def _build_pdf_tab(self, parent: Frame) -> None:
        sec = self._section(parent, "输入", "可多选：每页含多个价格签的 PDF · 支持拖入窗口 · 合并为一个总 PDF")
        frow = Frame(sec, bg=BG)
        frow.pack(fill="x")
        _btn(frow, "选择PDF文件", "secondary", self._pd_pick_files).pack(side="left")
        self.pd_analyze_btn = _btn(frow, "解析并检测", "secondary", self._pd_analyze, state=DISABLED)
        self.pd_analyze_btn.pack(side="left", padx=(10, 0))
        # 文件 chip 列表（每个文件一枚 chip，点 × 可单独删除）
        self.pd_chips = Frame(sec, bg=BG)
        self.pd_chips.pack(fill="x", pady=(6, 0))
        self.pd_chips_hint = Label(self.pd_chips, text="未选择文件（可多选）", font=FONT_LABEL, bg=BG, fg=INK_3, anchor="w")
        self.pd_chips_hint.pack(side="left")

        sec2 = self._section(parent, "检测结果", "按款式顺序 · 款式内先按颜色、再按尺码 XS→S→M→L→XL→XXL")
        self.pd_info = Label(sec2, text="请选择 PDF 并点击「解析并检测」", font=FONT_LABEL, bg=BG, fg=INK_2,
                             anchor="w", justify="left")
        self.pd_info.pack(fill="x")

        sec3 = self._section(parent, "输出参数", "页面 = 标签物理尺寸（参考样例 ≈ 4.54×7.54cm）")
        prow = Frame(sec3, bg=BG)
        prow.pack(fill="x")
        self.pd_per_page = IntVar(value=1)
        self.pd_margin = DoubleVar(value=0)
        self.pd_padding = DoubleVar(value=3)
        self.pd_label_w = DoubleVar(value=4.5)
        self.pd_label_h = DoubleVar(value=7.5)
        _field(prow, "标签宽 cm", self.pd_label_w).pack(side="left", padx=(0, 24))
        _field(prow, "标签高 cm", self.pd_label_h).pack(side="left", padx=(0, 24))
        _field(prow, "每页标签数", self.pd_per_page).pack(side="left", padx=(0, 24))
        _field(prow, "页边距 mm", self.pd_margin).pack(side="left", padx=(0, 24))
        _field(prow, "防切边 mm", self.pd_padding, "(默认3，不建议小于3)").pack(side="left")

        # v1.6：FSC 编号（每页标签左上角加盖的 FSC 标记编号，留空则不加盖）
        fsc_row = Frame(sec3, bg=BG)
        fsc_row.pack(fill="x", pady=(10, 0))
        self.pd_fsc_number = StringVar(value="C202518")
        _field(fsc_row, "FSC 编号", self.pd_fsc_number,
               "(默认 C202518，留空不加盖)").pack(side="left")

        arow = Frame(parent, bg=BG)
        arow.pack(fill="x", padx=28, pady=12)
        _btn(arow, "生成重排 PDF（合并）", "primary", self._pd_generate).pack(side="right")

        sec4 = self._section(parent, "输出信息", "")
        self.pd_status = ScrolledText(sec4, height=8, state=DISABLED, font=FONT_MONO,
                                     bd=0, bg=LINE_2, fg=INK, relief="flat", padx=10, pady=8)
        self.pd_status.pack(fill="x")

        self.pd_files: list = []   # 多个 PDF 路径

    def _pd_pick_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择 PDF 文件（可多选，也可直接拖入窗口）",
            filetypes=[("PDF", "*.pdf"), ("All", "*.*")],
        )
        if paths:
            self._pd_set_files(list(paths))

    def _pd_render_chips(self) -> None:
        self._render_chips(self.pd_files, self.pd_chips, self.pd_chips_hint,
                       self._pd_remove_files, self._pd_move_file)

    def _pd_remove_files(self, remain: list) -> None:
        self.pd_files = remain
        if self.pd_files:
            set_btn_state(self.pd_analyze_btn, NORMAL)
            self.status_var.set(f"剩余 {len(self.pd_files)} 个 PDF")
        else:
            set_btn_state(self.pd_analyze_btn, DISABLED)
            self.status_var.set("已移除全部 PDF 文件")
        self._pd_render_chips()

    def _pd_move_file(self, idx: int, delta: int) -> None:
        """交换 pd_files 中相邻两项，列表顺序即最终生成顺序。"""
        j = idx + delta
        if 0 <= j < len(self.pd_files):
            self.pd_files[idx], self.pd_files[j] = self.pd_files[j], self.pd_files[idx]
            self._pd_render_chips()
            self.status_var.set(
                f"顺序已调整：{os.path.basename(self.pd_files[j])} → 第 {j + 1} 位"
                f"（共 {len(self.pd_files)} 个）")

    # def _pd_set_files(self, paths: list) -> None:
    #     self.pd_files = list(paths)
    #     set_btn_state(self.pd_analyze_btn, NORMAL)
    #     self._pd_render_chips()
    #     _set_status(self.pd_status, f"已选择 {len(self.pd_files)} 个 PDF，请点击「解析并检测」")
    #     self.status_var.set(f"已选择 {len(self.pd_files)} 个 PDF")

    def _pd_set_files(self, paths: list, append: bool = True) -> None:
        """追加模式：新文件加到末尾，已存在的跳过（支持分多次拖入 / 选择）。"""
        cur = list(self.pd_files) if append else []
        dup = 0
        for p in paths:
            if p in cur:
                dup += 1
            else:
                cur.append(p)
        self.pd_files = cur

        if not self.pd_files:
            set_btn_state(self.pd_analyze_btn, DISABLED)
            self._pd_render_chips()
            self.status_var.set("未选择 PDF")
            return
        set_btn_state(self.pd_analyze_btn, NORMAL)
        self._pd_render_chips()
        tip = f"共 {len(self.pd_files)} 个 PDF" + (f"（跳过 {dup} 个重复）" if dup else "")
        _set_status(self.pd_status, tip + "，可拖拽 ↑↓ 调整顺序，然后点击「解析并检测」")
        self.status_var.set(tip)

    def _pd_analyze(self) -> None:
        if not self.pd_files:
            return
        try:
            import fitz
            lines = []
            grand = 0
            for path in self.pd_files:
                doc = fitz.open(path)
                total = 0
                per = []
                sizes = set()
                for i in range(len(doc)):
                    boxes = detect_stickers(doc[i])
                    total += len(boxes)
                    per.append(f"第 {i+1} 页：{len(boxes)} 枚")
                    for det in boxes:
                        w = det.padded.w / DETECT_SCALE
                        h = det.padded.h / DETECT_SCALE
                        sizes.add((round(w / 72 * 2.54, 2), round(h / 72 * 2.54, 2)))
                doc.close()
                grand += total
                sizes_str = ", ".join(f"{w}×{h}cm" for w, h in sorted(sizes)) or "-"
                lines.append(f"【{os.path.basename(path)}】共 {total} 枚 · 检测尺寸 {sizes_str}\n  " + "\n  ".join(per))
            self.pd_info.config(
                text=f"共 {len(self.pd_files)} 个款式 · 合计 {grand} 枚标签\n"
                     f"输出：1 个总 PDF（款式 → 颜色 → 尺码 XS→S→M→L→XL→XXL）", fg=INK)
            _set_status(self.pd_status, "\n".join(lines) +
                        "\n\n检测完成。所有款式将合并为一个 PDF，严格居中、尺寸完全一致。")
            self.status_var.set(f"检测完成：{len(self.pd_files)} 个款式 · {grand} 枚标签")
        except Exception as e:
            _set_status(self.pd_status, f"解析失败：{e}")
            self.status_var.set("解析失败")

    def _pd_generate(self) -> None:
        if not self.pd_files:
            messagebox.showwarning("提示", "请先选择 PDF 文件")
            return
        default_name = "价格签重排_汇总.pdf"
        out_path = filedialog.asksaveasfilename(
            title="保存为重排 PDF（所有款式合并为一个文件）",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=default_name,
        )
        if not out_path:
            return
        fsc_num = (self.pd_fsc_number.get() or "").strip()
        try:
            result = rearrange_pdfs(
                input_paths=self.pd_files,
                output_path=out_path,
                per_page=self.pd_per_page.get(),
                margin_mm=self.pd_margin.get(),
                padding_mm=self.pd_padding.get(),
                label_w_cm=self.pd_label_w.get(),
                label_h_cm=self.pd_label_h.get(),
                fsc_number=fsc_num,
            )
            lines = []
            for f in result["files"]:
                seq = " ｜ ".join(f"{g['color']}: {' '.join(g['sizes'])}" for g in f["color_groups"])
                lines.append(
                    f"✅ 【{f['style']}】{f['labels']} 枚\n"
                    f"   排列：{seq}"
                )
            msg = (
                f"已生成：{os.path.basename(out_path)}\n"
                f"合计 {result['labels']} 枚 · {result['pages']} 页 · "
                f"单枚 {result['label_w_cm']:.2f} × {result['label_h_cm']:.2f} cm\n"
                f"输出页 {result['page_w_cm']:.2f} × {result['page_h_cm']:.2f} cm\n"
                f"FSC 标记：{'已加盖 ' + fsc_num if fsc_num else '未加盖'}\n\n"
                + "\n".join(lines)
            )
            _set_status(self.pd_status, msg)
            self.status_var.set(f"已生成汇总 PDF：{os.path.basename(out_path)}（{result['labels']} 枚）")
        except Exception as e:
            _set_status(self.pd_status, f"生成失败：{e}")
            self.status_var.set("生成失败")


# ================================================================
# 入口
# ================================================================

def main() -> None:
    root = None
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()   # 支持「拖入导入」
    except Exception:
        root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
