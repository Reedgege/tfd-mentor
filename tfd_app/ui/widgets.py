# -*- coding: utf-8 -*-
"""
论文格式医生 · 中国版共享 UI 零件（严格照施工包 SVG 规格）
========================================================

所有零件都从 ``ui.theme`` 取色，禁止在零件里硬编码第二套样式。
可用零件（COMPONENT_MAP.md）：
  RoundButton / RoundCard / Badge / Stepper / InfoPanel / DropZone / Footer

按钮：主操作填 primary，次操作填 surface + primary 描边，删除类填 danger
（照 primary-button / secondary-button / danger-button.svg）。圆角取 RADIUS.control(8)。

零件自带文案统一取 ``COPY``（COPY_TABLE.md 第 5 节逐字照抄），不得在零件里另写措辞。
"""
from __future__ import annotations

import tkinter as tk
from . import theme
from .theme import RADIUS, SPACING, TYPE, get_theme


# ---------------------------------------------------------------------------
# 装修包零件权威文案（COPY_TABLE.md 第 5 节逐字照抄，唯一来源）
# ---------------------------------------------------------------------------
# 零件默认值一律从这里取，保证同一语义只有一套措辞；表里未列出的措辞
# 不在此新增（各页面自己的业务文案不放进这张表）。
COPY = {
    # 上传区（upload-dropzone.svg）
    "dropzone_title": "点击上传或拖拽文件到此处",
    # 选择框占位（select.svg）
    "select_placeholder": "请选择模板",
    # 套餐卡按钮（upgrade-modal.svg）
    "select": "选择",
    # 激活码输入框占位（upgrade-modal.svg）
    "input_code_ph": "输入激活码",
    # 复选框（checkbox-off.svg / checkbox-on.svg）
    "checkbox_off": "选择此项",
    "checkbox_on": "已选择",
    # 危险按钮（danger-button.svg）
    "danger": "删除 / 取消",
    # 禁用按钮（disabled-button.svg）
    "disabled": "暂不可用",
    # 状态（success.svg / processing.svg / error.svg）
    "processing": "正在处理，请稍候…",
    "done": "处理完成",
    "error": "处理失败，请检查文件",
    # 提示条（toast.svg）
    "toast_saved": "已保存 · 你的设置已更新",
}


def _rounded_rect(canvas, x1, y1, x2, y2, r, **kw):
    """在 canvas 上画圆角矩形（polygon 近似，tk 无原生 round-rect）。"""
    r = min(r, (x2 - x1) / 2, (y2 - y1) / 2)
    pts = [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]
    canvas.create_polygon(pts, smooth=True, **kw)


class RoundButton(tk.Frame):
    """圆角按钮（主 / 次 / 危险 / 禁用），API 兼容 ttk 按钮控件的 config(text/state/command)。

    危险态照 danger-button.svg（浅朱砂底 + 朱砂描边 + 朱砂字），
    用于删除类操作，不得拿主按钮伪装。
    """

    def __init__(self, master, text="", edition="student", style="primary",
                 font=None, height=SPACING["btn_h"], width=None, command=None, **kw):
        # width：显式像素宽（兼容 ttk 按钮控件的 width= 写法）；为 None 时按文字自适应。
        self._width_hint = width
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        self._style = style
        self._text = text
        self._command = command
        self._enabled = True
        self._hovered = False
        self._font = font or self._t.sans(TYPE["btn"], bold=True)
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0, height=height)
        self.cv.pack(fill="both", expand=True)
        self.cv.bind("<Configure>", lambda e: self._draw())
        self._bind_events()
        self._apply_width()
        self._draw()

    # ---- 宽度（避免 tk 画布默认 ~378px 撑爆按钮行）----
    def _apply_width(self):
        if self._width_hint:
            self.cv.configure(width=self._width_hint)
            return
        if self._text:
            try:
                f = tk.font.Font(font=self._font)
                px = int(f.measure(self._text))
            except Exception:
                px = len(self._text) * 14
            self.cv.configure(width=max(80, px + 36))
        else:
            self.cv.configure(width=80)

    # ---- 事件 ----
    def _bind_events(self):
        for w in (self, self.cv):
            w.bind("<Enter>", self._on_enter)
            w.bind("<Leave>", self._on_leave)
            w.bind("<Button-1>", self._on_click)

    def _on_enter(self, _e=None):
        if self._enabled:
            self._hovered = True
            self._draw()

    def _on_leave(self, _e=None):
        self._hovered = False
        self._draw()

    def _on_click(self, _e=None):
        if self._enabled and self._command:
            self._command()

    # ---- 配色 ----
    def _colors(self):
        t = self._t
        if not self._enabled:
            return t.disabled_fill, t.disabled_border, t.disabled_text
        if self._style == "primary":
            fill = t.primary_hover if self._hovered else t.primary
            return fill, fill, "#FFFFFF"
        if self._style == "danger":
            fill = t.danger_hover if self._hovered else t.danger_fill
            return fill, t.danger_border, t.danger_text
        # secondary
        fill = t.surface
        border = t.primary if self._hovered else t.secondary_border
        return fill, border, t.primary

    def _draw(self):
        self.cv.delete("all")
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        if w < 2 or h < 2:
            self.after(20, self._draw)
            return
        fill, border, fg = self._colors()
        _rounded_rect(self.cv, 1, 1, w - 1, h - 1, RADIUS["control"],
                      fill=fill, outline=border, width=2)
        self.cv.create_text(w / 2, h / 2, text=self._text, fill=fg,
                            font=self._font, anchor="center")

    # ---- 兼容 ttk 按钮控件的 config ----
    def config(self, **kw):
        if "text" in kw:
            self._text = kw.pop("text")
            self._apply_width()
        if "width" in kw:          # ttk 的显示宽度，显式指定则采用
            self._width_hint = kw.pop("width")
            self._apply_width()
        if "state" in kw:
            self._enabled = (kw.pop("state") != "disabled")
        if "command" in kw:
            self._command = kw.pop("command")
        if kw:
            super().config(**kw)
        self._draw()

    def set_text(self, text):
        self._text = text
        self._apply_width()
        self._draw()

    def set_style(self, style):
        """切换主 / 次按钮样式（交付方式三选一这类分段选择用，不另造第二套按钮）。"""
        self._style = style
        self._draw()


class RoundCard(tk.Frame):
    """圆角卡片：canvas 画圆角底，内层 Frame 放内容。"""

    def __init__(self, master, edition="student", fill=None, border=None,
                 radius=None, padx=SPACING["card_pad"], pady=SPACING["card_pad"],
                 corner=None, **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        self._fill = fill or self._t.surface
        self._border = border or self._t.border
        self._radius = radius or RADIUS["card"]
        self._padx = padx
        self._pady = pady
        # 圆角外那四个小角露出的是「页底色」，不是系统灰；默认取主题底色。
        self.corner = corner or self._t.bg
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0, bg=self.corner)
        self.cv.pack(fill="both", expand=True)
        self.inner = tk.Frame(self.cv, bg=self._fill)
        self._win = self.cv.create_window(self._padx, self._pady,
                                          window=self.inner, anchor="nw")
        self.cv.bind("<Configure>", lambda e: self._draw())
        self.inner.bind("<Configure>", lambda e: self._resize_inner())
        self.after(20, self._draw)

    def _resize_inner(self):
        w = self.cv.winfo_width()
        self.cv.coords(self._win, self._padx, self._pady)
        self.cv.itemconfig(self._win, width=max(2, w - self._padx * 2))

    def _draw(self):
        self.cv.delete("all")
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        if w < 2 or h < 2:
            self.after(20, self._draw)
            return
        _rounded_rect(self.cv, 1, 1, w - 1, h - 1, self._radius,
                      fill=self._fill, outline=self._border, width=2)
        self._win = self.cv.create_window(self._padx, self._pady,
                                          window=self.inner, anchor="nw")
        self._resize_inner()


class Badge(tk.Frame):
    """胶囊徽标：trial（蓝底）/ formal（绿灰底）。text 可变。"""

    def __init__(self, master, kind="trial", edition="student", text="", **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        self._kind = kind
        self._text = text
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0,
                            height=SPACING["badge_h"])
        self.cv.pack()                      # 不 fill：宽度由文字决定，避免 Canvas 默认 ~378px 撑成椭圆
        self.cv.bind("<Configure>", lambda e: self._draw())
        self._apply_width()
        self._draw()

    def _apply_width(self):
        """按文字测量宽度（含左右留白），避免 Canvas 默认 ~378px 撑成椭圆。"""
        try:
            f = tk.font.Font(font=self._t.sans(TYPE["body"], bold=True))
            w = max(64, int(f.measure(self._text)) + 28)
        except Exception:
            w = max(64, len(self._text) * 14 + 28)
        self.cv.configure(width=w)

    def _colors(self):
        t = self._t
        if self._kind == "formal":
            return t.formal_fill, t.formal_text
        return t.primary_soft, t.primary

    def _draw(self):
        self.cv.delete("all")
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        if w < 2:
            self.after(20, self._draw)
            return
        fill, fg = self._colors()
        _rounded_rect(self.cv, 1, 1, w - 1, h - 1, RADIUS["pill"], fill=fill,
                      outline=fill, width=2)
        self.cv.create_text(w / 2, h / 2, text=self._text, fill=fg,
                            font=self._t.sans(TYPE["body"], bold=True),
                            anchor="center")

    def set_text(self, text):
        self._text = text
        self._apply_width()
        self._draw()

    def set_kind(self, kind):
        """切换胶囊配色（trial 试用蓝 / formal 正式绿灰）—— 同一零件两种状态，不另造样式。"""
        self._kind = kind
        self._draw()


class Stepper(tk.Frame):
    """水平三步条（准备 / 检查 / 修正）。暴露 circles / titles / lines 列表，
    兼容既有 ``_refresh_wizard`` 的 config 调用。

    自适应：三列标题按自身宽度定宽（不换行、不截断），连接线吃掉剩余空间，
    说明行按列宽换行 —— 1100 最小窗口下右卡内宽约 483px，三列标题 + 两条
    连接线必须正好铺满，绝不把第三列挤出卡片被裁掉。
    """

    def __init__(self, master, steps, edition="student", **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        t = self._t
        self._step_circle = []
        self._step_title = []
        self._step_line = []
        self._step_desc = []
        n = len(steps)
        top = tk.Frame(self, bg=t.surface)
        top.pack(fill="x")
        for i, (label, desc) in enumerate(steps):
            col = tk.Frame(top, bg=t.surface)
            col.grid(row=0, column=2 * i, sticky="n")
            circ = tk.Label(col, text=str(i + 1), width=3, height=1,
                            relief="flat", highlightthickness=2,
                            highlightbackground=t.step_pending_border,
                            bg=t.step_pending_fill, fg=t.muted,
                            font=t.sans(TYPE["step_num"], bold=True))
            circ.pack()
            title = tk.Label(col, text=label, bg=t.surface, fg=t.muted,
                             font=t.sans(TYPE["card_title"], bold=True))
            title.pack(pady=(SPACING["hair"], 0))
            if desc:
                d = tk.Label(col, text=desc, bg=t.surface, fg=t.muted,
                             font=t.sans(TYPE["caption"]))
                d.pack(fill="x")
                self._step_desc.append(d)
            self._step_circle.append(circ)
            self._step_title.append(title)
            if i < n - 1:
                ln = tk.Frame(top, width=SPACING["step_line"], height=2,
                              bg=t.stepper_line)
                ln.grid(row=0, column=2 * i + 1, sticky="")
                self._step_line.append(ln)
            else:
                self._step_line.append(None)
        for i in range(n):
            top.grid_columnconfigure(2 * i, weight=1)   # 三列等宽展开，内容居中
        top.bind("<Configure>", lambda _e: self._reflow())
        self.after(20, self._reflow)

    def _reflow(self):
        """按实测宽度收紧：标题单行定宽；说明行按列宽换行；连接线吃掉余量。"""
        w = self.winfo_width()
        if w < 2:
            return
        n = len(self._step_title)
        if not n:
            return
        gaps = max(1, n - 1)
        # 先等分三列（扣掉连接线的最小呼吸位），再把余量补给连接线。
        # 标题只在「一列放不下」时才折行 —— 1100 最小窗下右卡内宽 483，
        # 最长的一步名（8 字 ≈ 198px）会折成两行，其余两步仍单行。
        line_w = SPACING["sm"] if n > 1 else 0
        colw = max(80, (w - line_w * gaps) // n)
        slack = w - colw * n - line_w * gaps
        if n > 1 and slack > 0:
            line_w = min(SPACING["step_line"], line_w + slack // gaps)
        # 连接线必须落在圆点的垂直中线上：三列已 sticky="n" 顶部对齐，线框若仍按
        # 整格居中就会被两行标题撑下去、压在文字上（1100 / 1280 实测正是如此）。
        cy = max(0, self._step_circle[0].winfo_reqheight() // 2 - 1)
        for ln in self._step_line:
            if ln is not None:
                ln.configure(width=line_w)
                ln.grid_configure(sticky="n", pady=(cy, 0))
        for t in self._step_title:
            t.configure(wraplength=self._balanced_wrap(t, colw))
        for d in self._step_desc:
            d.configure(wraplength=self._balanced_wrap(d, colw))

    @staticmethod
    def _balanced_wrap(label, colw):
        """折行宽度：能单行就单行；必须折行时收到「刚好还是这么多行」的最窄值，
        让各行字数尽量均匀 —— 避免出现「…论文文 / 件」这种单字孤行。"""
        txt = label.cget("text")
        if not txt:
            return colw
        try:
            f = tk.font.Font(font=label.cget("font"))
            widths = [f.measure(ch) for ch in txt]
        except Exception:
            return colw

        def lines_at(limit):
            n, cur = 1, 0
            for cw in widths:
                if cur + cw > limit and cur > 0:
                    n += 1
                    cur = cw
                else:
                    cur += cw
            return n

        target = lines_at(colw)
        if target <= 1:
            return colw
        best = colw
        for limit in range(colw, 20, -1):
            if lines_at(limit) > target:
                break
            best = limit
        return best

    @property
    def circles(self):
        return self._step_circle

    @property
    def titles(self):
        return self._step_title

    @property
    def lines(self):
        return self._step_line


class InfoPanel(tk.Frame):
    """「如何工作」信息面板（info-panel.svg）。"""

    def __init__(self, master, edition="student", title="如何工作",
                 line1="", line2="", **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        t = self._t
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0)
        self.cv.pack(fill="both", expand=True)
        self.inner = tk.Frame(self.cv, bg=t.info_fill)
        self._win = self.cv.create_window(0, 0, window=self.inner, anchor="nw")
        self.cv.bind("<Configure>", lambda e: self._draw())
        self.title_label = tk.Label(self.inner, text=title, bg=t.info_fill,
                                    fg=t.stepper_label_active,
                                    font=t.sans(TYPE["card_title"], bold=True))
        self.title_label.pack(anchor="w", padx=SPACING["card_pad"],
                              pady=(SPACING["sm"], 2))
        # 面板高度跟随内容：Canvas 默认请求尺寸是 378x265，不跟着内容走会撑出一大块空面。
        self.inner.bind("<Configure>", lambda e: self._fit())
        self.line1_lbl = None
        self.line2_lbl = None
        if line1:
            self.line1_lbl = tk.Label(self.inner, text=line1, bg=t.info_fill, fg=t.muted,
                                      font=t.sans(TYPE["body"]))
            self.line1_lbl.pack(anchor="w", padx=SPACING["card_pad"])
        if line2:
            self.line2_lbl = tk.Label(self.inner, text=line2, bg=t.info_fill, fg=t.muted,
                                      font=t.sans(TYPE["body"]))
            self.line2_lbl.pack(anchor="w", padx=SPACING["card_pad"],
                                pady=(2, SPACING["sm"]))
        self.after(20, self._draw)

    def _fit(self):
        """把画布高度收成内容实际高度（否则底部会留一大片空白画布）。"""
        h = self.inner.winfo_reqheight()
        if h > 2 and abs(self.cv.winfo_reqheight() - h) > 1:
            self.cv.configure(height=h)

    def _draw(self):
        self.cv.delete("all")
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        if w < 2 or h < 2:
            self.after(20, self._draw)
            return
        _rounded_rect(self.cv, 1, 1, w - 1, h - 1, RADIUS["card"],
                      fill=self._t.info_fill, outline=self._t.info_border, width=2)
        self._win = self.cv.create_window(0, 0, window=self.inner, anchor="nw")
        self.cv.coords(self._win, 0, 0)
        self.cv.itemconfig(self._win, width=w)


class DropZone(tk.Frame):
    """上传区（upload-dropzone.svg）：虚线圆角边框 + 箭头 + 标题 + 副标题。

    ``height``：画布高度，默认取 SPACING["dropzone_h"]（主界面实际值）；
    内部元素按实测高度等比摆放，不会挤出边框。
    ``title`` 默认取装修包权威文案 ``COPY["dropzone_title"]``。
    """

    def __init__(self, master, edition="student", title=None,
                 subtitle="", command=None, height=SPACING["dropzone_h"], **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        self._command = command
        self._height = height
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0, height=height)
        self.cv.pack(fill="both", expand=True)
        self.cv.bind("<Configure>", lambda e: self._draw())
        self.cv.bind("<Button-1>", lambda e: command and command())
        self._title = title if title is not None else COPY["dropzone_title"]
        self._subtitle = subtitle
        self.after(20, self._draw)

    def _draw(self):
        self.cv.delete("all")
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        if w < 2 or h < 2:
            self.after(20, self._draw)
            return
        t = self._t
        _rounded_rect(self.cv, 2, 2, w - 2, h - 2, RADIUS["card"],
                      fill=t.dropzone_fill, outline=t.dropzone_border, width=2,
                      dash=(6, 4))
        # 上箭头 + 两行文字：h ≥ 140 时按零件 SVG 的比例逐点摆放（h=150 与原件一致）；
        # h 较小时（主界面取 SPACING["dropzone_h"]=78）改用固定像素行距并整组垂直居中 ——
        # 字号不随画布缩小，纯按比例摆会让箭头 / 标题 / 副标题叠在一起（实测过）。
        cx = w / 2
        # 上传区主文案 ≤ 12pt（TYPE_SCALE.md 第 2 节硬约束）：参考稿里它比区块标题(14)小。
        title_font = t.sans(TYPE["card_title"], bold=True)
        sub_font = t.sans(TYPE["caption"])
        if h >= 140:
            arrow_top, arrow_tip = h * 0.267, h * 0.493
            title_y, sub_y = h * 0.613, h * 0.787
        else:
            # 紧凑分布：字号不随画布缩小，纯按比例摆会让箭头 / 标题 / 副标题相压。
            # 以 SPACING["dropzone_h"]（主界面实测值）的安全落位为 1 倍：
            # 箭头 3→21、主文案中线 36、副文案中线 57（68px 内不挤不裁）。
            k = h / float(SPACING["dropzone_h"])
            arrow_top = 3.0 * k
            arrow_tip = 21.0 * k
            title_y = 36.0 * k
            sub_y = 57.0 * k
        self.cv.create_line(cx, arrow_top, cx, arrow_tip, fill=t.primary, width=4,
                            arrow=tk.LAST, arrowshape=(10, 12, 5))
        self.cv.create_text(cx, title_y, text=self._title, fill=t.primary,
                            font=title_font, anchor="center")
        if self._subtitle:
            self.cv.create_text(cx, sub_y, text=self._subtitle, fill=t.muted,
                                font=sub_font, anchor="center")

    def set_subtitle(self, text):
        self._subtitle = text
        self._draw()


class SelectButton(tk.Frame):
    """选择框（select.svg）：圆角矩形 + 左对齐文字 + 右侧向下箭头，点击触发选择。"""

    def __init__(self, master, edition="student", placeholder=None,
                 command=None, height=SPACING["field_h"], **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        self._command = command
        self._placeholder = (placeholder if placeholder is not None
                             else COPY["select_placeholder"])
        self._text = self._placeholder
        self._height = height
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0, height=height)
        self.cv.pack(fill="both", expand=True)
        self.cv.bind("<Configure>", lambda e: self._draw())
        self.cv.bind("<Button-1>", lambda e: command and command())
        self.after(20, self._draw)

    def _draw(self):
        self.cv.delete("all")
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        if w < 2 or h < 2:
            self.after(20, self._draw)
            return
        t = self._t
        _rounded_rect(self.cv, 1, 1, w - 1, h - 1, RADIUS["control"],
                      fill=t.surface, outline=t.select_border, width=2)
        # 文字左对齐，垂直居中
        self.cv.create_text(18, h / 2, text=self._text, fill=t.select_ph,
                            font=t.sans(TYPE["body"]), anchor="w")
        # 向下箭头
        ax = w - 26
        ay = h / 2 - 4
        self.cv.create_line(ax, ay, ax + 7, ay + 8, ax + 14, ay,
                            fill=t.primary, width=2, smooth=False)

    def set_text(self, text):
        self._text = text if text else self._placeholder
        self._draw()


class Footer(tk.Frame):
    """页脚（footer.svg + COPY_TABLE 第 4 节）：

    第一行 左 ``✉ hi@reedskill.com`` / 中 ``专注学术 · 让格式更专业`` / 右 ``🌐 reedskill.com``；
    第二行 客服微信 · 官方公众号 · 本机处理声明 · 版权。
    微信 / 公众号只写栏目名，不写任何 ID（严禁虚构）。
    """

    def __init__(self, master, edition="student",
                 mail="hi@reedskill.com", site="reedskill.com", **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        t = self._t
        sep = tk.Frame(self, bg=t.border)
        sep.pack(fill="x")
        row = tk.Frame(self, bg=t.bg)
        row.pack(fill="x", pady=(SPACING["sm"], 0))
        tk.Label(row, text="✉ %s" % mail, bg=t.bg, fg=t.muted,
                 font=t.sans(TYPE["footer"])).pack(side="left")
        tk.Label(row, text="专注学术 · 让格式更专业", bg=t.bg, fg=t.muted,
                 font=t.sans(TYPE["footer"])).pack(side="left", expand=True)
        tk.Label(row, text="🌐 %s" % site, bg=t.bg, fg=t.muted,
                 font=t.sans(TYPE["footer"])).pack(side="right")
        row2 = tk.Frame(self, bg=t.bg)
        row2.pack(fill="x", pady=(SPACING["xs"], 0))
        tk.Label(row2, text="客服微信 · 官方公众号", bg=t.bg, fg=t.muted,
                 font=t.sans(TYPE["footer"])).pack(side="left")
        tk.Label(row2, text="© 2026 论文格式医生", bg=t.bg, fg=t.muted,
                 font=t.sans(TYPE["footer"])).pack(side="right")
        tk.Label(row2, text="本机处理 · 文件不会上传任何服务器", bg=t.bg,
                 fg=t.muted, font=t.sans(TYPE["footer"])).pack(side="right",
                                                              padx=(0, SPACING["lg"]))
