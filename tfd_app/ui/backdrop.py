# -*- coding: utf-8 -*-
"""
背景层：把施工包正式背景图（STUDENT_BACKGROUND.png / ADVISOR_BACKGROUND.png）
铺在窗口最底层。禁止使用完整 UI 预览图作背景（AGENT_MASTER_INSTRUCTION.md）。

实现：在 root 最底层放一个 Canvas（place 铺满），用 Tk 8.6 原生 PhotoImage 把背景图
按 cover 缩放后绘上（整数 zoom 放大 + 居中裁切，纯标准库、不引第三方依赖）。

另外提供两个给主界面用的小能力：
  - ``pin()``        页面滚动时把背景图钉回视口左上角（纹理不跟着内容跑）
  - ``sample_hex()`` 取图上某块区域的平均纸色（压在纸上的 widget 用它做底色，不露方底）
"""
from __future__ import annotations

import math
import os
import tkinter as tk

from . import theme
from .theme import get_theme
from .. import assetpath

_HERE = os.path.dirname(os.path.abspath(__file__))
_ASSETS = os.path.normpath(os.path.join(_HERE, "..", "assets"))


def _backdrop_path(edition):
    """背景图路径：优先走 assetpath（打包后布局会变），再退回源码树老路径。"""
    name = "student_background.png" if edition == "student" else "advisor_background.png"
    found = assetpath.find_asset(name)
    if found:
        return found
    p = os.path.join(_ASSETS, name)
    return p if os.path.isfile(p) else None


class Backdrop:
    """背景画布，创建后须调用 attach(root) 并随窗口 resize 重绘。"""

    def __init__(self, edition="student"):
        self._t = get_theme(edition)
        self._edition = edition
        self._path = _backdrop_path(edition)
        self._tk = None
        self._scaled = None          # 缓存按倍数放大后的图，避免每次 resize 重算
        self._scaled_key = None
        self._item = None            # 背景图在画布上的 item id
        self._img_pos = None         # 背景图摆放的基准坐标（未滚动时）
        self.canvas = None

    # --------------------------------------------------------------- 装载
    def _ensure_image(self):
        """按需加载原图（与画布尺寸无关，取色也要用）。取不到返回 None。"""
        if self._tk is None and self._path:
            try:
                self._tk = tk.PhotoImage(file=self._path)
            except Exception:
                self._tk = None
        return self._tk

    def _geom(self, w, h):
        """cover 规则：返回 (画布上图片的宽, 高, 整数放大倍数)；无图返回 None。"""
        img = self._ensure_image()
        if img is None:
            return None
        iw, ih = img.width(), img.height()
        if iw < 1 or ih < 1:
            return None
        if w > iw or h > ih:
            factor = max(1, math.ceil(max(w / iw, h / ih)))
            return iw * factor, ih * factor, factor
        return iw, ih, 1

    def attach(self, root):
        self.canvas = tk.Canvas(root, highlightthickness=0, bd=0, bg=self._t.bg)
        # place 铺满；用 Tk window lower 命令把背景沉到最底层
        self.canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self.canvas.tk.call("lower", self.canvas._w)
        root.configure(bg=self._t.bg)
        self.canvas.bind("<Configure>", lambda e: self._redraw())
        self._redraw()

    # --------------------------------------------------------------- 绘制
    def _redraw(self):
        if not self.canvas:
            return
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        if w < 2 or h < 2:
            self.canvas.after(20, self._redraw)
            return
        # 只删自己画的背景项 —— 主界面把整页内容（卡片 / 文字 / 线条）也画在这张画布上，
        # delete("all") 会把它们一起清掉（2026-09-12 实测踩到）。
        self.canvas.delete("bd_bg")
        self._item = None
        self._img_pos = None
        geom = self._geom(w, h)
        if geom is None:
            # 找不到背景图：退回主题底色（不报错、不影响启动）
            self.canvas.create_rectangle(0, 0, w, h, fill=self._t.bg, outline=self._t.bg,
                                         tags="bd_bg")
            return
        bw, bh, factor = geom
        key = (bw, bh)
        if factor > 1:
            if self._scaled is not None and self._scaled_key == key:
                img = self._scaled
            else:
                img = self._tk.zoom(factor, factor)
                self._scaled, self._scaled_key = img, key
        else:
            img = self._tk
            self._scaled, self._scaled_key = None, key
        left = (bw - w) // 2
        top = (bh - h) // 2
        # 图放在负偏移、由画布裁掉溢出部分 → 居中 cover（纯标准库没有 crop）
        self._img_pos = (-left, -top)
        self._item = self.canvas.create_image(-left, -top, anchor="nw", image=img,
                                              tags="bd_bg")
        self.canvas.tag_lower("bd_bg")      # 背景永远压在最底层

    def pin(self):
        """把背景图重新钉在当前视口左上角（页面滚动时纹理不跟着内容跑）。"""
        if self.canvas is None or self._item is None or self._img_pos is None:
            return
        try:
            self.canvas.coords(self._item, self._img_pos[0],
                               self._img_pos[1] + self.canvas.canvasy(0))
            self.canvas.tag_lower(self._item)
        except Exception:
            pass

    # --------------------------------------------------------------- 取色
    def sample_hex(self, win_w, win_h, x, y, w, h):
        """取宣纸图上对应窗口区域 (x, y, w, h) 的平均色，返回 '#RRGGBB'。

        用途：压在背景图上的品牌区 / 印章是 widget，没有真正的透明；用同处的实际纸色
        当底色，视觉上就融进背景，不会露出一块突兀的方底。
        纯标准库（PhotoImage.get 逐点取色求平均）；取不到就回退主题底色。
        """
        img = self._ensure_image()
        try:
            if img is None or win_w < 2 or win_h < 2:
                return self._t.bg
            geom = self._geom(win_w, win_h)
            if geom is None:
                return self._t.bg
            bw, bh, factor = geom
            ox, oy = (bw - win_w) // 2, (bh - win_h) // 2
            iw, ih = img.width(), img.height()
            acc = [0, 0, 0]
            n = 0
            for j in range(5):                      # 7x5 个采样点足够稳，也不拖慢启动
                yy = int(min(max(((y + oy) + h * j / 4.0) / factor, 0), ih - 1))
                for i in range(7):
                    xx = int(min(max(((x + ox) + w * i / 6.0) / factor, 0), iw - 1))
                    px = img.get(xx, yy)
                    if isinstance(px, str):
                        px = px.replace(",", " ").split()
                    acc = [a + int(b) for a, b in zip(acc, px)]
                    n += 1
            if not n:
                return self._t.bg
            return "#%02X%02X%02X" % tuple(a // n for a in acc)
        except Exception:
            return self._t.bg