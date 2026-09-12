# -*- coding: utf-8 -*-
"""
论文格式医生 · 中国版弹窗（照 upgrade-modal.svg / confirm-modal.svg）
=============================================================

ModalShell  ：弹窗壳（14 圆角、象牙面、细边、单一主操作 + 安静取消）
ConfirmModal：通用确认（取消 / 确认）
（UpgradeModal / 套餐卡片已按老板 2026-09-12 要求移除：升级入口改为「简单小程序码 + 文字」。）

全部从 ui.theme 取色，禁止第二套样式。
"""
from __future__ import annotations

import tkinter as tk
from . import theme
from .theme import get_theme, RADIUS, SPACING, TYPE
from .widgets import COPY, RoundButton, _rounded_rect


class ModalShell(tk.Toplevel):
    """带圆角画布的弹窗基底。subclass 在 self.body 里放内容。"""

    def __init__(self, parent, edition="student", title="", subtitle="",
                 width=620, height=300, **kw):
        super().__init__(parent, **kw)
        self._t = get_theme(edition)
        self._width = width
        self._height = height
        # 父窗口可见时才设 transient：启动时 root 被 withdraw，子窗若 transient 到
        # 已隐藏的父窗会被 Tk 级联隐藏/不出现，导致启动激活窗卡死。模态抓取仍保留。
        if parent is not None and parent.winfo_viewable():
            self.transient(parent)
        self.grab_set()
        self.overrideredirect(True)
        self.configure(bg=self._t.modal_border)
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0,
                            width=width, height=height)
        self.cv.pack(fill="both", expand=True)
        self.inner = tk.Frame(self.cv, bg=self._t.modal_fill)
        self._win = self.cv.create_window(0, 0, window=self.inner, anchor="nw")
        self.cv.bind("<Configure>", lambda e: self._draw())
        self.after(20, self._draw)
        # 标题区
        self.head = tk.Frame(self.inner, bg=self._t.modal_fill)
        self.head.pack(fill="x", padx=40, pady=(34, 6))
        if title:
            tk.Label(self.head, text=title, bg=self._t.modal_fill,
                     fg=self._t.modal_title,
                     font=self._t.serif(TYPE["section_title"], bold=True)).pack(anchor="w")
        if subtitle:
            tk.Label(self.head, text=subtitle, bg=self._t.modal_fill,
                     fg=self._t.modal_sub,
                     font=self._t.sans(TYPE["caption"])).pack(anchor="w", pady=(4, 0))
        self.body = tk.Frame(self.inner, bg=self._t.modal_fill)
        self.body.pack(fill="both", expand=True, padx=40, pady=(10, 10))
        # Esc 关闭（除升级弹窗外都允许）
        self.bind("<Escape>", lambda e: self._safe_close())

    def _draw(self):
        self.cv.delete("all")
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        if w < 2 or h < 2:
            self.after(20, self._draw)
            return
        _rounded_rect(self.cv, 1, 1, w - 1, h - 1, RADIUS["modal"],
                      fill=self._t.modal_fill, outline=self._t.modal_border, width=2)
        self._win = self.cv.create_window(0, 0, window=self.inner, anchor="nw")
        self.cv.coords(self._win, 0, 0)
        self.cv.itemconfig(self._win, width=w)

    def _safe_close(self):
        try:
            self.destroy()
        except Exception:
            pass

    def center_on(self, parent):
        self.update_idletasks()
        pw = parent.winfo_rootx()
        ph = parent.winfo_rooty()
        w = self._width
        h = self._height
        # 高 DPI 下（本机 150%：tk scaling≈2.0）字形比 96 DPI 大约 1.5 倍，固定
        # 760x560 会把套餐卡 / 激活区 / 弹窗页脚裁掉（实测「永久」卡与页脚都看不见）。
        # 按内容实际需要长高长宽，上限贴住屏幕，保证弹窗里的文案一个不少。
        try:
            w = max(w, self.inner.winfo_reqwidth() + 6)
            h = max(h, self.inner.winfo_reqheight() + 6)
            w = min(w, self.winfo_screenwidth() - 40)
            h = min(h, self.winfo_screenheight() - 60)
        except Exception:
            pass
        x = pw + (parent.winfo_width() - w) // 2
        y = ph + (parent.winfo_height() - h) // 2
        # 屏幕内约束：四边都不越界。否则弹窗底部（按钮 / 页脚）会被屏幕裁掉，
        # 实测 1280x720 @150% 时「升级正式版」弹窗底部按钮看不见。
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        x = max(0, min(x, sw - w - 10))
        y = max(0, min(y, sh - h - 10))
        self.geometry("%dx%d+%d+%d" % (w, h, x, y))


class ConfirmModal(ModalShell):
    """通用确认：body 放提示文字，底部 取消 / 确认。result 经 callback 回传。"""

    def __init__(self, parent, edition="student", title="确认操作",
                 message="请确认是否继续当前操作。", on_confirm=None,
                 confirm_text="确认", cancel_text="取消", **kw):
        super().__init__(parent, edition=edition, title=title, width=620, height=300)
        tk.Label(self.body, text=message, bg=self._t.modal_fill, fg=self._t.modal_sub,
                 font=self._t.sans(TYPE["body"]), wraplength=520, justify="left",
                 anchor="w").pack(anchor="w", pady=(8, 24))
        btns = tk.Frame(self.body, bg=self._t.modal_fill)
        btns.pack(fill="x", pady=(6, 10))
        cancel = RoundButton(btns, text=cancel_text, edition=edition, style="secondary",
                              height=SPACING["btn_h"],
                              font=self._t.sans(TYPE["btn_small"], bold=True),
                              command=self._safe_close)
        cancel.pack(side="left")
        ok = RoundButton(btns, text=confirm_text, edition=edition, style="primary",
                         height=SPACING["btn_h"],
                         font=self._t.sans(TYPE["btn"], bold=True),
                         command=lambda: (on_confirm and on_confirm(), self._safe_close()))
        ok.pack(side="right")
        self.center_on(parent)
