# -*- coding: utf-8 -*-
"""品牌标识（照 student-mark / advisor-mark / *-seal.svg 精确绘制）。"""
from __future__ import annotations

import os
import tkinter as tk
import tkinter.font as tkfont
from . import theme
from .. import assetpath
from .theme import get_theme, SPACING, TYPE

# 做旧传统印章 PNG（抄作业：形状/色/竖排两字照设计文件；优化：边缘侵蚀+印泥斑驳+楷体）。
# 运行时走 assetpath.find_asset，打包后也能找到。
_SEAL_SS = 3


def build_brand(master, edition="student", bg=None):
    """品牌区：左侧 logo 字形 + 右侧「论文格式医生」标题 + 副标题。返回 Frame。

    副标题固定为视觉稿的 ``让论文格式更简单``：零件 SVG（mark / advisor-mark）里自带
    的那行「版本副标题」按 COPY_TABLE.md 第 1 节的注释**不用于主界面顶部**，
    主界面一律以视觉稿为准，故这里不再按 edition 分叉。

    ``bg``：本区底色（默认主题底色）。主界面把它设成该处宣纸背景的实测纸色 ——
    widget 没有真正的透明，压在背景图上的一小块方底要靠「取同处纸色」才不显突兀。
    """
    t = get_theme(edition)
    bg = bg or t.bg
    f = tk.Frame(master, bg=bg)
    # logo 字形（照 mark.svg 的 path，stroke=primary 4px）
    # 画布按 SPACING["mark_h"]（TYPE_SCALE.md 第 2 节：标记得 34～38px）反推，
    # 字形在画布内居中，四周留出描边宽度，不会裁掉任何一笔。
    logo = tk.Canvas(f, width=SPACING["mark_w"] + 8, height=SPACING["mark_h"] + 8,
                     bg=bg, highlightthickness=0)
    logo.pack(side="left", padx=(0, SPACING["md"]))
    _draw_logo(logo, t.primary)
    txt = tk.Frame(f, bg=bg)
    txt.pack(side="left", anchor="w")
    tk.Label(txt, text="论文格式医生", bg=bg, fg=t.navy,
             font=t.serif(TYPE["page_title"], bold=True)).pack(anchor="w")
    tk.Label(txt, text="让论文格式更简单", bg=bg, fg=t.muted,
             font=t.sans(TYPE["subtitle"])).pack(anchor="w", pady=(2, 0))
    return f


# mark.svg 的原始路径点（x,y 交替；三段：顶部菱形书签 / 左侧竖边+底弧 / 中竖）
_MARK_SVG = (25, 24, 55, 14, 85, 24, 55, 34, 25, 24,
             34, 31, 34, 51, 76, 51, 76, 31,
             55, 34, 55, 62)


def _draw_logo(canvas, color):
    """把 _MARK_SVG 等比缩放到 SPACING["mark_h"] 高并居中画进画布。"""
    xs, ys = _MARK_SVG[0::2], _MARK_SVG[1::2]
    scale = float(SPACING["mark_h"]) / float(max(ys) - min(ys))
    cw, ch = canvas.winfo_reqwidth(), canvas.winfo_reqheight()
    ox = (cw - (max(xs) - min(xs)) * scale) / 2.0 - min(xs) * scale
    oy = (ch - (max(ys) - min(ys)) * scale) / 2.0 - min(ys) * scale
    pts = []
    for i in range(0, len(_MARK_SVG), 2):
        pts += [_MARK_SVG[i] * scale + ox, _MARK_SVG[i + 1] * scale + oy]
    canvas.create_line(pts[0], pts[1], pts[2], pts[3], pts[4], pts[5],
                       pts[6], pts[7], pts[0], pts[1],
                       fill=color, width=4, joinstyle="round")
    canvas.create_line(pts[8], pts[9], pts[10], pts[11], pts[12], pts[13],
                       pts[14], pts[15], pts[8], pts[9],
                       fill=color, width=4, joinstyle="round")
    canvas.create_line(pts[16], pts[17], pts[18], pts[19],
                       fill=color, width=4, joinstyle="round")


def build_seal(master, edition="student", bg=None, scale=1.0):
    """朱砂印章：做旧传统风（边缘侵蚀 + 印泥斑驳 + 楷体），落盘 PNG 贴图。

    抄作业（硬规矩）：竖圆角矩形、朱砂 #B54D43、竖排「学/生」「导/师」、字色 #FFF8ED、
    尺寸由 TYPE["seal"]=12pt 反推。优化（老板授权）：硬边→自然破损边、满红底→印泥斑驳、
    宋体→楷体（simkai.ttf）。``scale`` 仅影响透明留白区，印章本体固定 12pt 体量。
    """
    t = get_theme(edition)
    chars = ("学", "生") if edition == "student" else ("导", "师")
    fnt = tkfont.Font(font=t.serif(TYPE["seal"], bold=True))
    cw = max(8, fnt.measure(chars[0]))
    lh = max(8, fnt.metrics("linespace"))
    pad_x = max(3, int(round(cw * 0.30)))
    gap = max(1, int(round(lh * 0.12)))
    pad_y = max(3, int(round(lh * 0.20)))
    x1, y1 = 7 * scale, 5 * scale
    w = float(cw + 2 * pad_x) * scale
    h = float(2 * lh + gap + 2 * pad_y) * scale
    canvas = tk.Canvas(master, width=int(round(x1 + w + 7 * scale)),
                       height=int(round(y1 + h)), bg=bg or t.bg,
                       highlightthickness=0)
    png = assetpath.find_asset("seal_%s.png" % edition)
    if png:
        try:
            img = tk.PhotoImage(file=png).subsample(_SEAL_SS)
            canvas.create_image(x1 + w / 2.0, y1 + h / 2.0, image=img, anchor="center")
            canvas._seal_img = img          # 防止 PhotoImage 被 GC
            return canvas
        except Exception:
            pass
    # 兜底矢量绘制（资源缺失时仍可显示，不应触发）
    rr = min(w * 0.26, h / 2)
    pts = [x1 + rr, y1, x1 + w - rr, y1, x1 + w, y1, x1 + w, y1 + rr,
           x1 + w, y1 + h - rr, x1 + w, y1 + h, x1 + w - rr, y1 + h,
           x1 + rr, y1 + h, x1, y1 + h, x1, y1 + h - rr, x1, y1 + rr,
           x1, y1]
    canvas.create_polygon(pts, smooth=True, fill=t.seal, outline=t.seal)
    cx = x1 + w / 2
    cy1 = y1 + (pad_y + lh * 0.5) * scale
    cy2 = y1 + (pad_y + lh + gap + lh * 0.5) * scale
    canvas.create_text(cx, cy1, text=chars[0], fill="#FFF8ED",
                       font=t.serif(TYPE["seal"], bold=True), anchor="center")
    canvas.create_text(cx, cy2, text=chars[1], fill="#FFF8ED",
                       font=t.serif(TYPE["seal"], bold=True), anchor="center")
    return canvas
