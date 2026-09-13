# -*- coding: utf-8 -*-
"""
论文格式医生 · 导师版 原生 Tkinter 界面（导师检查学生论文、生成批注副本）
==============================================
面向客户的简洁流程，不展示任何执行细节；一屏放得下，无需滚动：

  一 · 选择文件        —— 批量导入论文（必选）+ 学校模板（可选）
  二 · 按步骤操作      —— ① 提取学校模板要求 → ② 论文格式检查 → ③ 按学校要求一键修正
  三 · 处理状态        —— 只显示友好状态与结果确认，无原始日志

交互约定：
  - 「提取学校模板要求」完成后，弹出“学校模板要求”确认页（关键项可修改），
    客户确认后才生效；放弃则本次不使用画像。
  - 确认页字段先用“批注要求”，没有批注则回退到“样式定义”填充（字体/字号/
    行距/缩进/页边距/对齐等），保证提取结果不空、可核对可修改。
  - 「按学校要求一键修正」导出地址由客户选择；完成后弹确认页，
    并附检查报告 + 修改报告。
  - .doc / WPS .wps 一律先自动转 .docx 再处理（进程内转换，无黑框、不弹窗），
    原文件绝不被修改。
"""
import os
import sys
import json
import time
import queue
import tempfile
import hashlib
import threading
import subprocess
import shutil
import webbrowser

# 把 tfd_app 加入搜索路径（打包成 exe 后 sys.path 已含，但开发态下保险）
HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
CORE_DIR = os.path.join(HERE, "core")
if CORE_DIR not in sys.path:
    sys.path.insert(0, CORE_DIR)

from . import trial       # v1.3.56：试用计数（机器码绑定，1 次）——相对导入，PyInstaller 才收集
from . import watermark   # v1.3.56：试用水印（页眉页脚+正文穿插）
from . import assetpath   # v1.1.1：运行时资源多候选路径解析（打包后布局会变）
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter.font as tkfont

from . import engine, license

# Excel 修改明细（零第三方依赖，手写 xlsx）；CORE_DIR 兜底确保可导入
try:
    from xlsx_report import write_change_report_xlsx
except ImportError:
    sys.path.insert(0, CORE_DIR)
    from xlsx_report import write_change_report_xlsx
try:
    from group_report import aggregate, extract_issues_from_markdown, classify_issue
except ImportError:
    sys.path.insert(0, CORE_DIR)
    from group_report import aggregate, extract_issues_from_markdown, classify_issue


# 资源路径统一走 assetpath：源码运行与打包后的布局不一致，写死一条路径会在
# 打包版静默读不到（v1.1.0 就是「二维码不显示、水印图丢失」的真事故）。
# 找不到时退回旧路径，保持既有 `os.path.isfile(...)` 守卫语义（False → 不显示）。
ICON = assetpath.find_asset("icon.png") or os.path.join(HERE, "assets", "icon.png")
QRCODE = assetpath.find_asset("qrcode.png") or os.path.join(HERE, "assets", "qrcode.png")
MINIAPP_QRCODE = (assetpath.find_asset("miniapp_qrcode.png")
                  or os.path.join(HERE, "assets", "miniapp_qrcode.png"))
MINIAPP_NAME = "芦苇论文格式"
# 品牌 / 客服（文案统一来源，避免散落硬编码）
WECHAT_NAME = "芦苇不熬夜"
WECHAT_ID = "reedskill"
ABOUT_MAIL = "hi@reedskill.com"
HELP_HINT = "\n\n遇到问题？看帮助或关注公众号【%s】留言。" % WECHAT_NAME
OFFICIAL_SITE = "https://reedskill.com"        # 官网（大本营）；给人看的省略 https，代码里打开用全址
OFFICIAL_SITE_TEXT = "🌐 reedskill.com"   # 页脚 / 官网入口文案（COPY_TABLE 第 4 节）

# ---------------------------------------------------------------------------
# 配色：严格照施工包 03_SHARED/tokens/DESIGN_TOKENS.json 的 advisor 段
# 唯一来源 = ui.theme 的 advisor 调色板；禁止在此写死第二套颜色 / 字体。
# ---------------------------------------------------------------------------
from .ui import get_theme, backdrop, widgets, brand, SPACING, TYPE   # 共享 UI 零件（edition="advisor" 取色）
from .ui.modal import ModalShell, ConfirmModal
_ADV = get_theme("advisor")        # 导师版 = advisor 调色板
PAPER   = _ADV.bg          # 宣纸底（advisor 暖米白）
PANEL   = _ADV.surface     # 面板（亮卡面）
INK     = _ADV.navy        # 主文字（暖褐黑）—— 标题 / 重点
BODY    = "#444444"        # 正文（灰）—— 描述性文字
MUTED   = _ADV.muted       # 次要文字 / 页脚（浅灰）
LINE    = _ADV.border      # 细线
ACCENT  = _ADV.primary     # 墨青（章节条 / 强调）
CINNABAR= _ADV.seal        # 印章红（主按钮 / 身份章）
CINNABAR_D = "#93433B"     # 印章红深（主操作 hover/active，官方 seal 的深一档）
OKC     = _ADV.success_dot # 完成（墨绿点）
ERRC    = _ADV.error_text  # 出错（朱红）
RUN     = ACCENT           # 处理中（墨青）

# 字体：标题 / 章节 / 状态用楷体（KaiTi，文艺调性），正文 / 按钮 / 页脚用微软雅黑（清晰）。
# 字号唯一来源 = ui.theme 的 TYPE 表（照 _ui_ref/TYPE_SCALE.md 逐字取值）—— 这里只做
# 「语义名 → 字体族 + 加粗」的转发，严禁再写第二套字号字面量（新旧字号混排正是"突兀"的成因）。
# 运行时由 _init_fonts() 变成 tkfont.Font 命名对象（name -> Font），任何 widget / ttk style
# 引用 F_XXX 都取到权威值；字号不随窗口缩放（TYPE_SCALE.md 第 4 节：不许用缩放函数糊）。
_FONT_BASE = {
    "F_TITLE":      ("KaiTi", TYPE["page_title"], "bold"),      # 页面大标题（关于 / 帮助 / 激活窗）
    "F_SUB":        ("Microsoft YaHei", TYPE["subtitle"]),      # 副标题
    "F_HDR":        ("KaiTi", TYPE["section_title"], "bold"),   # 区块标题（楷体）
    "F_BODY":       ("Microsoft YaHei", TYPE["body"]),          # 正文 / 步骤说明
    "F_SMALL":      ("Microsoft YaHei", TYPE["caption"]),       # 次按钮 / 次要说明
    "F_SMALL_B":    ("Microsoft YaHei", TYPE["body"], "bold"),  # 加粗正文（字段名 / 提示框标题）
    "F_BTN":        ("Microsoft YaHei", TYPE["btn"], "bold"),   # 主按钮
    "F_STAT":       ("KaiTi", TYPE["body"]),                    # 状态文字（楷体）
    "F_SUBTITLE":   ("Microsoft YaHei", TYPE["card_title"]),    # 卡片标题 / 元信息
    "F_CARD_HDR":   ("KaiTi", TYPE["section_title"], "bold"),   # 区块标题（文件选择 / 处理步骤）
    "F_FOOT":       ("Microsoft YaHei", TYPE["footer"]),        # 状态栏 / 页脚 / 说明（最小层级）
    "F_MONO":       ("Consolas", TYPE["mono"]),                 # 机器码 / 离线码（等宽）
    "F_DIALOG_TITLE": ("KaiTi", TYPE["section_title"], "bold"), # 弹窗标题（楷体）
    "F_ICON":       ("KaiTi", TYPE["seal"], "bold"),            # 印章图标（论 / 模，楷体朱砂）
}
from .buildinfo import APP_VERSION   # 版本号唯一来源（与 VERSION 文件的一致性由 tests 钉住）

# v1.0.31：绿色 zip 版由软件自建桌面快捷方式（win32com 已内置，客户零依赖、零黑框）。
APP_SHORTCUT_NAME = "论文格式医生·导师版"   # 桌面快捷方式显示名
SHORTCUT_FLAG_TAG = "ThesisFormatDoctorMentor"  # %APPDATA% 下询问标记目录（区分学生/导师版）


def _is_frozen_exe():
    """打包后的主程序才显示“桌面图标”入口 / 首启询问；
    开发态（python.exe/pythonw.exe）与 mac/Linux 一律不建 Windows 快捷方式。"""
    base = os.path.basename(sys.executable or "").lower()
    return base.endswith(".exe") and not (base.startswith("python")
                                          or base.startswith("pythonw"))

  # 与 VERSION 文件保持同步（状态栏显示用）
_FONTS = {}      # name -> tkfont.Font（字号一律取自 TYPE，不随窗口缩放）

def _init_fonts(root):
    """在 root 创建后调用：把 F_XXX 全局名替换为 Font 命名对象（字号 = TYPE 权威值）。"""
    import platform as _platform
    # 跨平台中文字体兜底：Windows 自带楷体/雅黑；mac/Linux 无则映射到系统 CJK 字体，
    # 避免 tkinter 静默回退到无中文的默认字体出现方块（tofu）。Windows 上字体存在，走原值。
    _fams = set(tkfont.families())
    _KAITI = {"Darwin": "STKaiti", "Linux": "Noto Serif CJK SC"}
    _HEITI = {"Darwin": "PingFang SC", "Linux": "Noto Sans CJK SC"}
    def _resolve(fam):
        if fam in _fams:
            return fam
        cand = (_KAITI if fam == "KaiTi" else _HEITI).get(_platform.system())
        if cand and cand in _fams:
            return cand
        return fam
    for name, spec in _FONT_BASE.items():
        kw = {"family": _resolve(spec[0]), "size": spec[1]}
        if len(spec) > 2:
            kw["weight"] = spec[2]
        f = tkfont.Font(root=root, **kw)
        _FONTS[name] = f
        globals()[name] = f

# 页边距字段：编辑厘米值时同步写 twips（top/bottom/left/right），兼容两套读取方
_MARGIN_TWIPS = {
    ("spec", "page", "top_cm"): "top",
    ("spec", "page", "bottom_cm"): "bottom",
    ("spec", "page", "left_cm"): "left",
    ("spec", "page", "right_cm"): "right",
}
_TWIPS_PER_CM = 567.0   # 1 厘米 ≈ 567 twips

ALIGN_DISPLAY = {"center": "居中", "left": "左对齐", "right": "右对齐", "both": "两端对齐"}
ALIGN_CODE = {v: k for k, v in ALIGN_DISPLAY.items()}

# 确认页可修改项：(中文标签, [候选取值路径…], 控件类型)
# 取值优先级：批注要求(spec.*) → 样式定义(profile 根的 body/headings/headingStyles/page)
# 字段取值路径：levels 段最完整（样式定义挖全要素 + 批注覆盖），spec 兜底，根级再兜底。
# levels 键：zh_font/sz(半磅)/align/line_val/indent_chars/before_pt/after_pt/bold 等。
EDIT_FIELDS = [
    ("正文字体",     [("levels", "body", "zh_font"), ("spec", "body", "zh_font"),
                     ("spec", "body", "font"), ("body", "font")], "text"),
    # v1.3.49：字号候选把 sz（半磅数字）提到 size（中文字号名）之前——
    # 此前 size 优先导致表单显示"小四"、客户改数字写回 size 字段而引擎只读 sz，
    # 客户改字号不生效（自查发现的逻辑 bug）。
    ("正文字号",     [("levels", "body", "sz"), ("levels", "body", "size"),
                     ("spec", "body", "sz"), ("spec", "body", "size")], "text"),
    ("正文行距(磅)", [("levels", "body", "line_val"), ("spec", "body", "line_val"),
                     ("body", "line_val"), ("body", "line")], "text"),
    ("首行缩进(字符)", [("levels", "body", "indent_chars"), ("spec", "body", "indent_chars"),
                       ("body", "indent_chars"), ("body", "firstLineChars")], "text"),
    ("一级标题字体", [("levels", "1", "zh_font"), ("spec", "h1", "zh_font"),
                     ("spec", "h1", "font"), ("headings", "1", "font")], "text"),
    ("一级标题字号", [("levels", "1", "sz"), ("levels", "1", "size"),
                     ("spec", "h1", "sz"), ("spec", "h1", "size")], "text"),
    ("一级标题对齐", [("levels", "1", "align"), ("spec", "h1", "align")], "align"),
    ("二级标题字体", [("levels", "2", "zh_font"), ("spec", "h2", "zh_font"),
                     ("spec", "h2", "font"), ("headings", "2", "font")], "text"),
    ("二级标题字号", [("levels", "2", "sz"), ("levels", "2", "size"),
                     ("spec", "h2", "sz"), ("spec", "h2", "size")], "text"),
    ("三级标题字体", [("levels", "3", "zh_font"), ("spec", "h3", "zh_font"),
                     ("spec", "h3", "font"), ("headings", "3", "font")], "text"),
    ("三级标题字号", [("levels", "3", "sz"), ("levels", "3", "size"),
                     ("spec", "h3", "sz"), ("spec", "h3", "size")], "text"),
    ("页边距 上(厘米)", [("spec", "page", "top_cm"), ("page", "top")], "text"),
    ("页边距 下(厘米)", [("spec", "page", "bottom_cm"), ("page", "bottom")], "text"),
    ("页边距 左(厘米)", [("spec", "page", "left_cm"), ("page", "left")], "text"),
    ("页边距 右(厘米)", [("spec", "page", "right_cm"), ("page", "right")], "text"),
    ("参考文献格式", [("levels", "reference"), ("spec", "reference"),
                     ("refExample",), ("reference",)], "ref"),
    # v1.3.78：结构页（摘要/致谢/附录）确认项——模板画像已拆「标题/正文」双 spec，
    # 弹窗中可逐项核对与修改；写回 levels.xxx_title / levels.xxx 供引擎套用。
    ("摘要标题字体", [("levels", "abstract_title", "zh_font"),
                     ("spec", "abstract_title", "zh_font"),
                     ("levels", "abstract", "zh_font")], "text"),
    ("摘要标题字号", [("levels", "abstract_title", "sz"), ("levels", "abstract_title", "size"),
                     ("spec", "abstract_title", "sz"), ("levels", "abstract", "sz")], "text"),
    ("摘要正文字体", [("levels", "abstract", "zh_font"), ("spec", "abstract", "zh_font")], "text"),
    ("摘要正文字号", [("levels", "abstract", "sz"), ("levels", "abstract", "size"),
                     ("spec", "abstract", "sz")], "text"),
    ("摘要正文行距(磅)", [("levels", "abstract", "line_val"), ("spec", "abstract", "line_val")], "text"),
    ("致谢标题字体", [("levels", "ack_title", "zh_font"), ("spec", "ack_title", "zh_font"),
                    ("levels", "ack", "zh_font")], "text"),
    ("致谢标题字号", [("levels", "ack_title", "sz"), ("levels", "ack_title", "size"),
                    ("spec", "ack_title", "sz"), ("levels", "ack", "sz")], "text"),
    ("致谢正文字体", [("levels", "ack", "zh_font"), ("spec", "ack", "zh_font")], "text"),
    ("附录标题字体", [("levels", "appendix_title", "zh_font"), ("spec", "appendix_title", "zh_font"),
                    ("levels", "appendix", "zh_font")], "text"),
    ("附录标题字号", [("levels", "appendix_title", "sz"), ("levels", "appendix_title", "size"),
                    ("spec", "appendix_title", "sz"), ("levels", "appendix", "sz")], "text"),
    # v1.3.79：附录正文一律不自动修改（各校差异极大），故不提供"附录正文字体"确认项，
    # 避免客户修改后不生效造成误导；附录标题处会以批注说明正文请自行对照学校要求处理。
]


def _base_no_ext(path):
    return os.path.splitext(path)[0]


def _deep_get(d, path):
    for k in path:
        if not isinstance(d, dict) or k not in d:
            return None
        d = d[k]
    return d


def _deep_set(d, path, value):
    for k in path[:-1]:
        if not isinstance(d.get(k), dict):
            d[k] = {}
        d = d[k]
    d[path[-1]] = value


_CN_SIZE_PT = {  # 中文字号名 → 磅（通用印刷规范，非学校特定值）
    "初号": 42, "小初": 36, "一号": 26, "小一": 24, "二号": 22, "小二": 18,
    "三号": 16, "小三": 15, "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5, "七号": 5.5, "八号": 5,
}


def _cn_size_to_pt(text):
    """把客户输入的字号转成磅：支持数字（如 12）与中文字号名（如 小四=12）。

    v1.3.49：确认弹窗字号输入此前只认数字，客户填"小四"会被当非法忽略。
    """
    s = (text or "").strip()
    if not s:
        return None
    try:
        return float(s)
    except (TypeError, ValueError):
        pass
    if s in _CN_SIZE_PT:
        return _CN_SIZE_PT[s]
    # 容忍"小四号""小五号"等写法：先匹配"小X"（小四=12 优先于 四号=14），再匹配普通字号名
    for name in ("小初", "小一", "小二", "小三", "小四", "小五", "小六", "小七"):
        if name in s:
            return _CN_SIZE_PT[name]
    for name, pt in _CN_SIZE_PT.items():
        if name in s:
            return pt
    return None


def _field_value(profile, candidates):
    """取确认页字段值：批注要求优先，样式定义兜底；含单位换算。"""
    for path in candidates:
        v = _deep_get(profile, path)
        if v in (None, ""):
            continue
        s = str(v)
        # levels/spec 里的 sz 是半磅（OOXML 单位）→ 显示为磅
        if path[-1] == "sz":
            try:
                return str(int(float(v)) // 2)
            except (ValueError, TypeError):
                return s
        # 样式定义里的页边距是 twips → 换算成厘米显示
        if len(path) == 2 and path[0] == "page" and path[1] in ("top", "bottom", "left", "right"):
            try:
                return "%.2f" % (float(v) / _TWIPS_PER_CM)
            except (ValueError, TypeError):
                return s
        # 样式定义里的正文行距是 twips → 换算成磅显示
        if tuple(path) == ("body", "line"):
            try:
                return str(int(float(v) / 20.0))
            except (ValueError, TypeError):
                return s
        return s
    return ""


def _profile_summary(profile):
    """把格式画像渲染成客户看得懂的“学校模板要求”文本行。

    批注要求(spec.*)优先，样式定义(profile 根)兜底，保证提取结果不空。
    """
    lines = []
    spec = profile.get("spec") or {}
    levels = profile.get("levels") or {}

    # 页面
    page = spec.get("page") or {}
    if not page.get("top_cm") and not page.get("top"):
        page = profile.get("page") or {}
    if page.get("top_cm") is not None:
        lines.append("· 页边距：上 %s 下 %s 左 %s 右 %s（厘米）" % (
            page.get("top_cm"), page.get("bottom_cm"),
            page.get("left_cm"), page.get("right_cm")))
    elif page.get("top") is not None:
        try:
            t = "%.2f" % (float(page["top"]) / _TWIPS_PER_CM)
            b = "%.2f" % (float(page["bottom"]) / _TWIPS_PER_CM)
            l = "%.2f" % (float(page["left"]) / _TWIPS_PER_CM)
            r = "%.2f" % (float(page["right"]) / _TWIPS_PER_CM)
            lines.append("· 页边距：上 %s 下 %s 左 %s 右 %s（厘米）" % (t, b, l, r))
        except (ValueError, TypeError):
            lines.append("· 页边距：已提取")

    # 正文（levels 完整画像优先，spec/根级兜底）
    body = levels.get("body") or spec.get("body") or profile.get("body") or {}
    if body:
        parts = []
        font = body.get("zh_font") or body.get("font")
        if font:
            parts.append("字体 %s" % font)
        size = body.get("size")
        if size:
            parts.append("字号 %s" % size)
        elif body.get("sz"):
            try:
                parts.append("字号 %s" % str(int(float(body["sz"])) // 2))
            except (ValueError, TypeError):
                pass
        indent = body.get("indent_chars") or body.get("firstLineChars")
        if indent:
            parts.append("首行缩进 %s 字符" % indent)
        line_val = body.get("line_val")
        if line_val is None and body.get("line"):
            try:
                line_val = int(float(body["line"]) / 20.0)
            except (ValueError, TypeError):
                line_val = None
        if line_val:
            parts.append("行距 %s 磅" % line_val)
        lines.append("· 正文：%s" % ("，".join(parts) if parts else "样式已提取"))

    # 各级标题（levels 优先，spec/headings 兜底）
    lv_keys = {"h1": "1", "h2": "2", "h3": "3"}
    for lv, name in (("h1", "一级标题"), ("h2", "二级标题"), ("h3", "三级标题")):
        h = (levels.get(lv_keys[lv]) or spec.get(lv)
             or profile.get("headings", {}).get(lv_keys[lv]) or {})
        if not h:
            continue
        parts = []
        font = h.get("zh_font") or h.get("font")
        if font:
            parts.append("字体 %s" % font)
        size = h.get("size")
        if size:
            parts.append("字号 %s" % size)
        elif h.get("sz"):
            try:
                parts.append("字号 %s" % str(int(float(h["sz"])) // 2))
            except (ValueError, TypeError):
                pass
        if h.get("align"):
            parts.append("对齐 %s" % ALIGN_DISPLAY.get(h["align"], h["align"]))
        if h.get("bold"):
            parts.append("加粗")
        lines.append("· %s：%s" % (name, "，".join(parts) if parts else "样式已提取"))

    # 其它分类（levels 优先，spec 兜底；dict 或文本都展示）
    for key, label in (("abstract", "摘要"), ("keywords", "关键词"), ("toc", "目录"),
                       ("reference", "参考文献"), ("title", "论文题目"),
                       ("table", "表格"), ("figure", "插图"), ("footnote", "脚注")):
        v = levels.get(key) or spec.get(key)
        if isinstance(v, dict) and v:
            s = "，".join("%s %s" % (k, val) for k, val in list(v.items())[:4])
            lines.append("· %s：%s" % (label, s))
        elif isinstance(v, str) and v.strip():
            lines.append("· %s：%s" % (label, v.strip()[:40]))

    if not lines:
        lines.append("（未从模板提取到明确的格式要求，将按通用规范处理。）")
    return lines


class _CanvasText:
    """把「宣纸页上的 canvas 文字项」包装成能 ``.config(text=/fg=)`` 的对象。

    页脚 / 状态栏必须画成 canvas 文字项 —— widget 自带底色，会把宣纸纹理盖成一块米白。
    而既有逻辑只会调 ``.config()``，这里做一层薄适配，业务逻辑一行都不用改。
    """

    def __init__(self, app, item):
        self._app = app
        self._cv = app._page
        self._item = item

    def config(self, text=None, fg=None, fill=None, **_kw):
        if text is not None:
            self._cv.itemconfig(self._item, text=text)
        color = fg if fg is not None else fill
        if color is not None:
            self._cv.itemconfig(self._item, fill=color)
        # 文字变了 → 状态栏是右对齐拼排的，跟着重排一次
        self._app._layout_status()
        return self

    configure = config


class App:
    def __init__(self, root):
        global _APP_REF
        self.root = root
        _APP_REF = self
        self.root.title("论文格式医生 · 导师版")
        # 导师版 = advisor 调色板 + 宣纸背景（严格照施工包 advisor 段）
        self.edition = "advisor"
        self._theme = get_theme("advisor")
        self._backdrop = backdrop.Backdrop(self.edition)
        try:
            self._backdrop.attach(self.root)
        except Exception:
            pass
        # v1.0.11：主区已可滚动，窗口不再需要靠"撑得巨大"来露出页脚。
        # 默认尺寸按屏幕自适应（不超过屏幕 86%×88%），小屏也能容纳；最小尺寸放宽，
        # 再小的窗口也只是主区内部滚动，页脚 / 状态栏始终可见。
        _sw = self.root.winfo_screenwidth()
        _sh = self.root.winfo_screenheight()
        # 新主界面按「宣纸页」排版（品牌区 116 + 卡片 + 页脚 126 + 状态栏 34）。
        # 这两个尺寸还要拿去给背景取色（品牌区 / 印章的底色要贴合该处纸色，见 _build_widgets）。
        self._win_w = min(1180, int(_sw * 0.86))
        self._win_h = min(880, int(_sh * 0.88))
        self.root.geometry("%dx%d" % (self._win_w, self._win_h))
        # 最小尺寸按新排版重定：1100 宽保证 50:50 等分后两卡各约 506px（按钮文字不换行），
        # 且顶部「品牌组｜右上入口行」两条带子之间仍留得出安全间距（实测至少 60px，不压叠）；
        # 690 高保证卡片区放得下左卡内容（实测约 418px），此时不出现滚动条。
        self.root.minsize(1100, 690)
        # v1.0.10：Windows 上、且屏幕分辨率 ≥ 1440×900 时启动即最大化，
        # 让页脚/状态栏在最大化窗口里绝对可见（避免用户拖到小窗口时把页脚裁掉）；
        # 1366×768 等小屏幕跳过，让用户保留窗口控制权（最大化后无关闭按钮风险）。
        # macOS/Linux 跳过（行为不同）。
        try:
            if sys.platform.startswith("win"):
                sw = self.root.winfo_screenwidth()
                sh = self.root.winfo_screenheight()
                if sw >= 1440 and sh >= 900:
                    self.root.state("zoomed")
        except tk.TclError:
            pass
        try:
            if os.path.isfile(ICON):
                # ⚠️ 必须把 PhotoImage 存成属性保留引用：直接写
                # `root.iconphoto(True, tk.PhotoImage(file=ICON))` 时，这个临时对象
                # 在语句结束后立刻被 CPython 回收，它的 __del__ 会连带把 Tk 里的图片
                # 删掉 —— 表现就是「换了图标却还是 Tk 默认羽毛图标」。2026-09-11 实测：
                # 不保留引用时 GC 后 Tk 的 image names 里查不到我们的图标。
                self._window_icon = tk.PhotoImage(file=ICON)
                self.root.iconphoto(True, self._window_icon)
        except Exception:
            self._window_icon = None

        # 随窗口缩放自适应：监听尺寸变化只重排坐标；字号恒为 TYPE 权威值，不随窗口缩放。
        self._last_w = None
        self.root.bind("<Configure>", self._on_resize)

        self.thesis_path = tk.StringVar()      # 兼容旧引用（取 thesis_paths[0] 的命名）
        self.thesis_paths = []                  # 批量：已选论文路径列表
        self.template_path = tk.StringVar()
        self.profile_path = tk.StringVar()
        self.author_var = tk.StringVar(value="论文格式医生·导师版")  # 批注署名（导师名）
        # 导师版交付方式：批注副本 / 一键修正 / ①+② 都要（三选一，默认只批注不改原稿）
        self._fix_mode = "annotate"
        self._fix_mode_var = tk.StringVar(value="annotate")  # 交付模式：annotate / fix / both
        self.status_var = tk.StringVar(value="请按步骤操作")
        self.running = False
        self._dialog_open = False   # 保存对话框打开期间防重复弹窗
        self._profile_confirmed = False  # 画像是否已被客户确认（检查/修正前弹出确认页）
        self._profile_abandoned = False  # v1.3.49：客户在确认页点"放弃"后，本次不再自动重新提取画像
        # 已保存文件记录：同一会话内再次保存到同一文件时弹“已保存过，是否再次保存”
        self._check_report_saved = None   # 第②步检查报告已保存路径
        self._fix_saved = set()           # 第③步修正产出文件已保存路径集合（兼容旧引用）
        self._fix_outs = {}               # 第③步修正完成后的临时产出：{mode: (dst, chk, rep)}
        # v1.0.15：第③步三个交付方式各自独立记录状态（互不覆盖）：
        #   _fix_phase[mode] = idle→fixed→saved；_exported[mode] = 该方式已保存的文件/文件夹。
        # 这样切到未导出的方式可正常生成导出；切回已导出的方式按钮显示「已完成」，
        # 点击提示"已导出过"，确认后可重新生成覆盖，杜绝静默重复导出。
        self._fix_phase = {}
        self._exported = {}
        self._errored = False
        self._msgs = []
        # 步骤名走装修包的横向步进器（circle + 标签），标签位只放短名；
        # 第③步的三种交付方式在右卡的「交付方式」区展开。
        self.step_defs = [("profile", "提取学校模板要求"),
                          ("check", "论文格式检查"),
                          ("fix", "一键格式修正")]
        # 每步的说明行（照 COPY_TABLE.md 第 3 节）
        self.step_descs = ["识别字号、页边距与格式规范",
                           "快速定位格式问题与待修正项",
                           "确认后生成符合规范的论文文件"]
        self.step_index = 0

        # 兜底：直接 ``App(root)``（不经 gui.main()）时 F_XXX 字体名尚未建立，界面会
        # NameError。正规启动路径 main() 已先调 _init_fonts；这里只在缺失时补一次，
        # 幂等、不改任何业务行为（仅为「构造 App 不抛异常」）。
        if globals().get("F_BODY") is None:
            _init_fonts(self.root)

        self._build_style()
        self._build_widgets()
        self._restore_template_lock()

        # v1.0.32：绿色 zip 版首次启动自动创建桌面快捷方式（静默，仅 Windows 正式版生效）
        self.root.after(800, self._maybe_auto_shortcut)

    # -------------------------------------------------- 窗口尺寸自适应
    def _on_resize(self, _evt=None):
        try:
            w = self.root.winfo_width()
        except Exception:
            return
        if w < 60:
            return
        if w == self._last_w:
            return
        self._last_w = w
        # 宽度变了 → 重排一次（背景取色 / 卡片宽度 / 文字位置全部跟着走）
        self.root.after(1, self._layout)

    # ------------------------------------------------------------- style
    def _build_style(self):
        """ttk 只留主题 + 进度条样式。

        主界面按钮一律用 ``ui.widgets.RoundButton``（装修包唯一按钮零件），
        ttk 不再定义自定义按钮样式 —— 同类元素只允许一套规范样式。
        """
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        # 进度条（处理中）取色同样来自 ui.theme，不写第二套色
        style.configure("Paper.Horizontal.TProgressbar",
                        background=_ADV.primary, troughcolor=_ADV.disabled_fill,
                        bordercolor=_ADV.disabled_fill,
                        lightcolor=_ADV.primary, darkcolor=_ADV.primary)

    # ------------------------------- 版式常量（唯一来源 = ui.theme 的 SPACING 表）
    # 字号缩小后这些固定像素高度必须同步收紧，否则会出现「字小了、块还是那么大」的空洞感。
    MX = SPACING["page_mx"]         # 页面左右留白
    HEAD_H = SPACING["head_h"]      # 顶部品牌区高度（mark 高 36 / 朱砂竖章随 12pt 字号）
    FOOT_H = SPACING["foot_h"]      # 页脚高度（细线 + 两行文字）
    STATUS_H = SPACING["status_h"]  # 状态栏高度
    CARD_GAP = SPACING["card_gap"]  # 两卡之间的一档留白
    CARD_PAD = SPACING["card_pad"]  # 卡片内边距（与 RoundCard 的 padx/pady 一致）

    def _build_widgets(self):
        """主界面 = 宣纸页 + 顶部品牌区 + 右上入口行 + 两张并排卡片 + 页脚 + 状态栏。

        版式基准 = 共享完整视觉稿
        ``_ui_ref/00_READ_FIRST/ORIGINAL_HANDOFF_FILES/01_UI_MOCKUP_REFERENCE.png``
        （导师版按同一版式装配，只换身份与专属功能）。顶部左侧是品牌组：线描 mark
        （``brand.build_brand``）＋**大标题**「论文格式医生」＋朱砂**竖**章「导师版」
        （``brand.build_seal``）＋副标题，四件成组靠左；右上是一条入口行：试用状态胶囊
        ＋「升级正式版 →」＋「客服微信｜官方公众号｜关于｜帮助」。
        （``02_ADVISOR/previews/ADVISOR_FINAL_LAYOUT_PREVIEW.png`` 是早期草图，无大标题、
        无入口行，不作为视觉依据。）

        为什么整页画在背景层 Canvas 上：tkinter 的 widget 一定有底色，用 Frame 铺满窗口
        就会把宣纸纹理盖成一块纯米白（实测旧界面正是如此）。canvas 的文字与细线没有底色，
        纹理才透得出来；只有卡片 / 按钮是真正的 widget —— 它们本来就该是「压在纸上的一块面」。
        全部元素坐标由 ``_layout`` 统一算，不再靠 pack 层层嵌套，从根上杜绝压叠 / 错位 / 截断。
        """
        self.root.configure(bg=PAPER)
        page = self._backdrop.canvas
        if page is None:                      # 背景层异常时的兜底，保证界面照常起
            page = tk.Canvas(self.root, bg=PAPER, highlightthickness=0)
            page.place(x=0, y=0, relwidth=1, relheight=1)
        self._page = page
        self._body_cv = page                  # 兼容旧引用名（页面滚动就是这个画布）
        self._laying_out = False
        self._in_yscroll = False
        self._sb_shown = False
        self._foot_top = 0
        self._status_y = 0

        # 页面滚动条：只在内容确实比窗口高时才 place 出来（验收要求「无多余滚动条」）
        self._page_sb = ttk.Scrollbar(self.root, orient="vertical", command=page.yview)
        page.configure(yscrollcommand=self._on_page_yscroll)
        page.bind("<Configure>", self._layout, add="+")   # add="+"：不覆盖背景层的重绘绑定

        # ---- 顶部品牌组（照视觉稿成组靠左）：线描 mark ＋ 大标题「论文格式医生」
        #      ＋ 朱砂竖章「导师版」＋ 副标题。四件装进同一个容器里并排，坐标只算容器一处，
        #      从结构上保证它们永远是一组、不会各自漂移或互相压叠。
        #      edition 恒为 "advisor" → build_brand / build_seal 渲染的就是导师版身份。
        head_bg = self._backdrop.sample_hex(self._win_w, self._win_h,
                                            self.MX, 10, 440, self.HEAD_H - 20)
        self._brand_group = tk.Frame(page, bg=head_bg)
        self._brand = brand.build_brand(self._brand_group, self.edition, bg=head_bg)
        self._brand.pack(side="left")
        self._seal = brand.build_seal(self._brand_group, self.edition, bg=head_bg,
                                      scale=1.0)
        self._seal.pack(side="left", padx=(SPACING["md"], 0))
        self._brand_win = page.create_window(0, 0, window=self._brand_group, anchor="nw")

        # ---- 右上入口行：试用状态胶囊 ＋「升级正式版 →」＋ 客服微信｜官方公众号｜关于｜帮助
        #      关于 / 帮助 接的仍是既有的 show_about / show_help，只是入口回到视觉稿的右上。
        self._trial_badge = widgets.Badge(page, kind="trial", edition=self.edition,
                                          text="试用版")
        self._badge_win = page.create_window(0, 0, window=self._trial_badge, anchor="e")
        self._upgrade_btn = widgets.RoundButton(
            page, edition=self.edition, text="升级正式版 →", style="primary",
            height=SPACING["field_h"], font=self._theme.sans(TYPE["btn"], bold=True),
            command=self._show_upgrade_qr)
        self._upgrade_win = page.create_window(0, 0, window=self._upgrade_btn, anchor="e")
        self._head_links = [
            self._paper_entry(page, "客服微信", self._show_service_qr, font=F_BODY),
            self._paper_entry(page, "官方公众号", self._show_official_qr, font=F_BODY),
            self._paper_entry(page, "关于", lambda: show_about(self.root), font=F_BODY),
            self._paper_entry(page, "帮助", lambda: show_help(self.root), font=F_BODY),
        ]
        self._head_seps = [
            page.create_text(0, 0, text="｜", anchor="e", fill=MUTED, font=F_BODY)
            for _ in range(len(self._head_links) - 1)]
        self._update_trial_badge()

        # ---- 主体：两张并排卡片（左 文件选择 / 右 处理步骤）
        self._left_card = widgets.RoundCard(page, edition=self.edition, bg=PAPER,
                                            corner=PAPER)
        self._lcard_win = page.create_window(0, 0, window=self._left_card, anchor="nw")
        self._right_card = widgets.RoundCard(page, edition=self.edition, bg=PAPER,
                                             corner=PAPER)
        self._rcard_win = page.create_window(0, 0, window=self._right_card, anchor="nw")
        self._build_left(self._left_card.inner)
        self._build_right(self._right_card.inner)
        # 卡片内容高度会变（选中文件后文件名变长、画像出现 / 删除…）：一变就重排整页，
        # 保证卡片永远装得下内容，不会出现半截被裁的字。
        for card in (self._left_card, self._right_card):
            card.inner.bind("<Configure>", lambda e: self._layout(), add="+")

        # ---- 页脚（COPY_TABLE 第 4 节）：
        #      上排 左「✉ 邮箱」· 中「专注学术 · 让格式更专业」· 右「🌐 官网」；
        #      下排 左「客服微信 · 官方公众号」· 中「© 2026 论文格式医生」
        #           · 右「本机处理 · 文件不会上传任何服务器」。
        #      （关于 · 帮助 按视觉稿留在右上入口行，页脚不再重复放一份）
        self._foot_rule = page.create_line(0, 0, 0, 0, fill=LINE)
        self._foot_mail = page.create_text(
            0, 0, anchor="w", fill=MUTED, font=F_FOOT, text="✉ " + ABOUT_MAIL)
        self._foot_slogan = page.create_text(
            0, 0, anchor="center", fill=MUTED, font=F_FOOT,
            text="专注学术 · 让格式更专业")
        self._foot_privacy = page.create_text(
            0, 0, anchor="e", fill=MUTED, font=F_FOOT,
            text="本机处理 · 文件不会上传任何服务器")
        self._foot_copy = page.create_text(
            0, 0, anchor="center", fill=MUTED, font=F_FOOT,
            text="© 2026 论文格式医生")
        # 页脚文字入口：canvas 文字没有底色，宣纸纹理不会被盖住（左→右排列）
        self._foot_cs = self._paper_entry(page, "客服微信", self._show_service_qr,
                                          anchor="w")
        self._foot_cs_dot = page.create_text(0, 0, text="·", anchor="w",
                                             fill=MUTED, font=F_FOOT)
        self._foot_off = self._paper_entry(page, "官方公众号", self._show_official_qr,
                                           anchor="w")
        self._foot_links = [
            self._paper_entry(page, OFFICIAL_SITE_TEXT,
                              lambda: webbrowser.open(OFFICIAL_SITE)),
        ]
        # v1.0.31：绿色 zip 版一键补建桌面快捷方式（win32com 已内置 exe，零外部依赖、零黑框）
        if sys.platform.startswith("win") and _is_frozen_exe():
            self._foot_links.append(
                self._paper_entry(page, "桌面图标", self._create_desktop_shortcut))

        # ---- 状态栏（页脚下方一行；同样是 canvas 文字，不盖纹理）
        self.bar_dot_item = page.create_text(0, 0, text="●", anchor="w", fill=OKC, font=F_FOOT)
        self.bar_left_item = page.create_text(0, 0, text="", anchor="w", fill=MUTED, font=F_FOOT)
        self.status_dot_item = page.create_text(0, 0, text="●", anchor="e", fill=MUTED, font=F_FOOT)
        self.status_lbl_item = page.create_text(0, 0, text="", anchor="e", fill=INK, font=F_FOOT)
        self.bar_right_item = page.create_text(0, 0, text="", anchor="e", fill=MUTED, font=F_FOOT)
        self.bar_dot = _CanvasText(self, self.bar_dot_item)
        self.bar_left = _CanvasText(self, self.bar_left_item)
        self.status_dot = _CanvasText(self, self.status_dot_item)
        self.status_lbl = _CanvasText(self, self.status_lbl_item)
        self.bar_right = _CanvasText(self, self.bar_right_item)
        self._status_rule = page.create_line(0, 0, 0, 0, fill=LINE)
        # 状态文字走 status_var（既有逻辑只 set 变量 / config 颜色），这里把变量接到画布上
        self.status_var.trace_add("write",
                                  lambda *_a: self.status_lbl.config(text=self.status_var.get()))
        self.status_lbl.config(text=self.status_var.get())

        def _on_wheel(e):
            # 焦点在可自行滚动的文本框内时不拦截，避免双重滚动
            if isinstance(e.widget, tk.Text):
                return
            page.yview_scroll(int(-1 * (e.delta / 120)), "units")
            self._backdrop.pin()

        self.root.bind_all("<MouseWheel>", _on_wheel)

        self._set_bar("idle")
        self.root.after(30, self._layout)

    # ------------------------------------------------------------ 版式计算
    def _paper_entry(self, canvas, text, cmd, anchor="e", font=None):
        """宣纸页上的文字入口（无底色 / 悬停变色 / 手型光标）。

        为什么要自己画：Label 一定有背景色，会把宣纸纹理盖出一块方底。
        ``font``：入口行字号（右上入口行 10pt / 页脚 9pt），默认取页脚字号。
        """
        item = canvas.create_text(0, 0, text=text, anchor=anchor, fill=ACCENT,
                                  font=font or F_FOOT)
        canvas.tag_bind(item, "<Button-1>", lambda e: cmd())
        canvas.tag_bind(item, "<Enter>",
                        lambda e: (canvas.itemconfig(item, fill=_ADV.primary_hover),
                                   canvas.configure(cursor="hand2")))
        canvas.tag_bind(item, "<Leave>",
                        lambda e: (canvas.itemconfig(item, fill=ACCENT),
                                   canvas.configure(cursor="")))
        return item

    def _text_w(self, item):
        """canvas 文字项的实测宽度（bbox 在窗口尚未显示时也算得出来）；空串算 0。"""
        try:
            if not self._page.itemcget(item, "text"):
                return 0
            bb = self._page.bbox(item)
            if bb:
                return bb[2] - bb[0]
        except Exception:
            pass
        return 72

    @staticmethod
    def _widget_w(win):
        """页面上某个 widget 的实测宽度（取不到时按 146 兜底，绝不因取宽失败把两串摆错位）。"""
        try:
            w = win.winfo_reqwidth()
            if w > 2:
                return w
        except Exception:
            pass
        return 146

    def _extra_height(self):
        """两张卡片内容实际需要的高度；比视口高时整页滚动，绝不裁掉内容。"""
        try:
            self.root.update_idletasks()
        except Exception:
            pass
        h = 0
        for card in (getattr(self, "_left_card", None), getattr(self, "_right_card", None)):
            if card is None:
                continue
            try:
                h = max(h, card.inner.winfo_reqheight() + 2 * self.CARD_PAD + 8)
            except Exception:
                continue
        return int(h)

    def _layout(self, _evt=None):
        """按窗口尺寸重排整页（页面元素全是 canvas 项，坐标必须自己算）。

        显式坐标的唯一目的：卡片 / 按钮 / 页脚 / 状态栏各自待在自己那一格里，
        从根上消掉旧版 pack 互相压叠、错位、被裁掉的实测硬伤。
        """
        page = getattr(self, "_page", None)
        if page is None or self._laying_out:
            return
        W, H = page.winfo_width(), page.winfo_height()
        if W < 2 or H < 2:
            return
        self._laying_out = True
        try:
            mx = self.MX
            colw = max(340, (W - 2 * mx - self.CARD_GAP) // 2)
            body_top = self.HEAD_H
            # 卡片区可用高度 = 窗口高 - 品牌区 - 页脚 - 状态栏 - 卡片与页脚间的一档留白
            avail = H - self.HEAD_H - self.FOOT_H - self.STATUS_H - 14
            body_h = max(240, avail, self._extra_height())
            # 顶部：品牌组（mark＋大标题＋朱砂竖章＋副标题）靠左垂直居中；
            #      右上入口行整条右对齐 —— 从右往左摆：帮助 … 关于 … 官方公众号 … 客服微信
            #      → 升级正式版按钮 → 试用状态胶囊。两条带子共用同一条中线，绝不重叠。
            page.coords(self._brand_win, mx,
                        max(0, (self.HEAD_H - self._brand_group.winfo_reqheight()) // 2))
            row_y = self.HEAD_H // 2
            row = []
            for i, item in enumerate(self._head_links):
                row.append(item)
                if i < len(self._head_seps):
                    row.append(self._head_seps[i])
            links_w = sum(self._text_w(i) + 4 for i in row)
            badge_w = self._widget_w(self._trial_badge)
            btn_w = self._widget_w(self._upgrade_btn)
            brand_w = self._brand_group.winfo_reqwidth()
            one_line = (mx + brand_w + 16 + badge_w + SPACING["sm"] + btn_w
                        + links_w <= W - mx)
            # 1100 最小窗实测：单行排不下（品牌组 385 + 入口行 719 + 左右留白 > 1100）。
            # 排不下时文字入口降一行、徽标与升级按钮留在上一行，两行都右对齐 ——
            # 五个入口一个不少，且绝不压到品牌组 / 大标题上。
            top_y = row_y if one_line else row_y - 12
            links_y = row_y if one_line else row_y + 14
            x = W - mx
            for item in reversed(row):
                page.coords(item, x, links_y)
                x -= self._text_w(item) + 4
            if not one_line:
                x = W - mx
            page.coords(self._upgrade_win, x, top_y)
            x -= btn_w + SPACING["sm"]
            page.coords(self._badge_win, x, top_y)
            # 主体两卡（严格等宽等高）
            page.coords(self._lcard_win, mx, body_top)
            page.itemconfig(self._lcard_win, width=colw, height=body_h)
            page.coords(self._rcard_win, mx + colw + self.CARD_GAP, body_top)
            page.itemconfig(self._rcard_win, width=colw, height=body_h)
            # 页脚：细线 + 上排（邮箱 / 口号 / 官网）+ 下排（客服入口 / 版权 / 本机声明）
            foot_top = body_top + body_h + 14
            self._foot_top = foot_top
            page.coords(self._foot_rule, mx, foot_top, W - mx, foot_top)
            y1 = foot_top + 20
            y2 = foot_top + 44
            page.coords(self._foot_mail, mx, y1)
            page.coords(self._foot_slogan, W // 2, y1)
            page.coords(self._foot_privacy, W - mx, y2)
            page.coords(self._foot_copy, W // 2, y2)
            x = W - mx
            for item in reversed(self._foot_links):
                page.coords(item, x, y1)
                x -= self._text_w(item) + 18
            x = mx
            page.coords(self._foot_cs, x, y2)
            x += self._text_w(self._foot_cs) + 6
            page.coords(self._foot_cs_dot, x, y2)
            x += self._text_w(self._foot_cs_dot) + 6
            page.coords(self._foot_off, x, y2)
            # 状态栏（贴页底一行）
            self._status_y = foot_top + self.FOOT_H + self.STATUS_H // 2 + 2
            page.coords(self._status_rule, mx, foot_top + self.FOOT_H,
                        W - mx, foot_top + self.FOOT_H)
            total = max(foot_top + self.FOOT_H + self.STATUS_H, H)
            page.configure(scrollregion=(0, 0, W, total))
            self._layout_status()
            # 滚动条有无：按可见比例算（不调 yview，避免滚动回调重入）
            self._on_page_yscroll(1.0 - min(1.0, float(H) / float(total)), 1.0)
        finally:
            self._laying_out = False

    def _layout_status(self):
        """状态栏一行：左边「就绪点 + 就绪文案」，右边「状态 + 提示」右对齐拼排。"""
        page = getattr(self, "_page", None)
        if page is None:
            return
        W = page.winfo_width()
        if W < 2:
            return
        y = self._status_y
        mx = self.MX
        page.coords(self.bar_dot_item, mx, y)
        page.coords(self.bar_left_item, mx + 16, y)
        page.coords(self.bar_right_item, W - mx, y)
        _w_right = self._text_w(self.bar_right_item)
        x = W - mx - (_w_right + 26 if _w_right else 0)
        page.coords(self.status_lbl_item, x, y)
        x -= self._text_w(self.status_lbl_item) + 6
        page.coords(self.status_dot_item, x, y)

    def _on_page_yscroll(self, first, last):
        """页面滚动回调：同步滚动条 + 把宣纸图钉回视口 + 按需显隐滚动条。"""
        if self._in_yscroll:
            return
        self._in_yscroll = True
        try:
            self._page_sb.set(first, last)
            self._backdrop.pin()
            need = (float(last) - float(first)) < 0.999
            if need and not self._sb_shown:
                self._page_sb.place(relx=1.0, x=-4, y=8, anchor="ne",
                                    height=max(80, self._page.winfo_height() - 16))
                self._page_sb.lift()
                self._sb_shown = True
            elif not need and self._sb_shown:
                self._page_sb.place_forget()
                self._sb_shown = False
        except Exception:
            pass
        finally:
            self._in_yscroll = False

    # ---------------------------------------------------------- left panel
    def _build_left(self, parent):
        """左卡：文件选择（上传区 / 批量导入 / 学校模板 / 记住模板 / 批注署名 / 模板说明）。

        文案一律照 COPY_TABLE.md 第 2 节。导师版专属的批量入口全部用装修包零件
        （upload-dropzone / folder-button / multi-file-button），接的仍是既有
        ``_pick_input``（多选）与 ``_pick_folder``（整文件夹，既有批量 Handler），
        **不新写任何处理逻辑**。
        """
        hdr = tk.Frame(parent, bg=PANEL)
        hdr.pack(fill="x", pady=(0, SPACING["xs"]))
        tk.Frame(hdr, bg=ACCENT, width=4, height=14).pack(side="left",
                                                          padx=(0, SPACING["sm"]))
        tk.Label(hdr, text="文件选择", bg=PANEL, fg=INK, font=F_CARD_HDR).pack(side="left")
        # 卡片右上注释（视觉稿的浅色说明文字，不再用徽章）
        tk.Label(hdr, text="支持多种文档格式，智能识别格式要求", bg=PANEL, fg=MUTED,
                 font=F_FOOT).pack(side="right")

        # 上传区（upload-dropzone 零件）：点击 / 拖拽选论文文件
        widgets.DropZone(
            parent, edition=self.edition, height=SPACING["dropzone_h"],
            title="点击上传或拖拽文件到此处",
            subtitle="支持 Word 文档（.docx、.docx）、WPS 格式（.wps）",
            command=self._pick_input).pack(fill="x")

        # 上传按钮（次按钮，secondary-button.svg）
        up_row = tk.Frame(parent, bg=PANEL)
        up_row.pack(fill="x", pady=(SPACING["xs"], SPACING["sm"]))
        widgets.RoundButton(up_row, text="选择文件", edition=self.edition,
                            style="secondary", height=SPACING["field_h"],
                            font=F_BTN,
                            command=self._pick_input).pack(side="left")

        # 批量导入（导师版专属）：整文件夹（folder-button，接既有 _pick_folder 批量
        # Handler）/ 多文件（multi-file-button，接既有 _pick_input）
        self._thesis_box, self._thesis_name, self._thesis_dot, self._thesis_lbl = self._file_row(
            parent, "论", "批量导入论文",
            "可选择文件夹，也可选择多个文件",
            "选择多个文件", self._pick_input,
            btn2_text="选择文件夹", cmd2=self._pick_folder)
        self._template_box, self._template_name, self._tpl_dot, self._tpl_lbl = self._file_row(
            parent, "模", "学校模板",
            "用于按学校要求检查／修正，更贴合要求", "选择模板", self._pick_template,
            badge_text="可选")

        # 批量处理进度（batch-progress.svg）：只在批量处理中出现，空闲时整块收起，
        # 不常驻空框。文案照零件：标题「批量处理进度」／计数「%d / %d」
        # （零件示例 12 / 18）／当前文件「正在处理%s」（零件示例 正在处理论文_12.docx）。
        self._batch_prog = tk.Frame(parent, bg=PANEL, highlightthickness=1,
                                    highlightbackground=LINE)
        _bp_top = tk.Frame(self._batch_prog, bg=PANEL)
        _bp_top.pack(fill="x", padx=SPACING["md"], pady=(SPACING["sm"], SPACING["xs"]))
        tk.Label(_bp_top, text="批量处理进度", bg=PANEL, fg=INK,
                 font=F_SMALL_B).pack(side="left")
        self._batch_count = tk.Label(_bp_top, text="", bg=PANEL, fg=MUTED,
                                     font=F_FOOT)
        self._batch_count.pack(side="right")
        self._batch_bar = ttk.Progressbar(self._batch_prog, mode="determinate",
                                          maximum=100, length=100,
                                          style="Paper.Horizontal.TProgressbar")
        self._batch_bar.pack(fill="x", padx=12)
        self._batch_file = tk.Label(self._batch_prog, text="", bg=PANEL, fg=MUTED,
                                    font=F_FOOT, anchor="w")
        self._batch_file.pack(fill="x", padx=SPACING["md"],
                              pady=(SPACING["xs"], SPACING["sm"]))

        # 记住模板：点一下把模板路径存本机，下次打开自动载入；再点取消（按钮绿=已记住）
        self._tpl_locked = False
        lock_row = tk.Frame(parent, bg=PANEL)
        lock_row.pack(fill="x", pady=(0, SPACING["xs"]))
        self._lock_btn = widgets.RoundButton(
            lock_row, text="🔒 记住此模板", edition=self.edition, style="secondary",
            height=SPACING["field_h"], font=F_SMALL, command=self._on_lock_click)
        self._lock_btn.pack(side="left")
        self._lock_tip = tk.Label(lock_row, text="", bg=PANEL, fg=OKC, font=F_SMALL_B)
        self._lock_tip.pack(side="left", padx=(SPACING["sm"], 0))

        # 批注署名：导师名，出现在 Word 批注气泡作者栏（默认"论文格式医生·导师版"）
        auth_row = tk.Frame(parent, bg=PANEL)
        auth_row.pack(fill="x", pady=(0, SPACING["xs"]))
        tk.Frame(auth_row, bg=ACCENT, width=4, height=14).pack(side="left",
                                                              padx=(0, SPACING["sm"]))
        tk.Label(auth_row, text="批注署名", bg=PANEL, fg=INK,
                 font=F_SUBTITLE).pack(side="left")
        tk.Entry(auth_row, textvariable=self.author_var, font=F_SMALL, relief="solid",
                 bd=1).pack(side="left", padx=(SPACING["sm"], 0), fill="x", expand=True)

        # 模板说明常驻（防止客户上传无批注模板造成误解，减少纠纷）：紧凑两行
        note = tk.Frame(parent, bg=PANEL)
        note.pack(fill="x")
        tk.Label(note, text="模板说明", bg=PANEL, fg=_ADV.primary,
                 font=F_SUBTITLE).pack(anchor="w")
        tk.Label(note, text="请优先使用学校官方模板（通常批注中写明了格式要求）；若模板无批注，将按模板格式定义／通用规范处理，可能与学校要求有出入。",
                 bg=PANEL, fg=MUTED, font=F_FOOT, anchor="w", justify="left",
                 wraplength=420).pack(anchor="w")

        # 画像状态（提取后由 _update_profile_box 显示在「学校模板」块下方）
        self.profile_box = tk.Frame(parent, bg=_ADV.success_fill,
                                    highlightthickness=1, highlightbackground=_ADV.success_border)
        tk.Label(self.profile_box, text="●", bg=_ADV.success_fill, fg=OKC,
                 font=F_FOOT).pack(side="left", padx=(SPACING["sm"], SPACING["xs"]),
                                   pady=SPACING["xs"])
        self.profile_info_var = tk.StringVar(value="已载入格式画像")
        tk.Label(self.profile_box, textvariable=self.profile_info_var, bg=_ADV.success_fill,
                 fg=_ADV.success_text, font=F_FOOT).pack(side="left", fill="x", expand=True)
        widgets.RoundButton(self.profile_box, text="删除 / 取消", edition=self.edition,
                            style="danger", height=SPACING["btn_h_sm"], font=F_SMALL,
                            command=self._clear_profile).pack(
                                side="right", padx=SPACING["sm"], pady=SPACING["xs"])
        self._update_profile_box()

        # 卡片底部标语（视觉稿）
        tk.Label(parent, text="一键规范格式，专注论文内容", bg=PANEL, fg=MUTED,
                 font=F_FOOT).pack(anchor="center", pady=(SPACING["sm"], 0))

    @staticmethod
    def _clip(s, n=24):
        """中文按字符截断，超长用省略号；用于文件名/模板名显示，避免撑坏布局。"""
        s = s or ""
        return s if len(s) <= n else s[:n - 1] + "…"

    def _file_row(self, parent, icon, title, desc, btn_text, cmd,
                  btn2_text=None, cmd2=None, badge_text=None):
        """一个「文件选择块」：图标 + 标题 + 选择按钮（首行）/ 状态 + 文件名（次行）。

        返回 (外框, 文件名标签, 状态点, 状态文字)：后三个是既有逻辑
        （_update_thesis_status / _pick_template / _restore_template_lock / _reset_wizard）
        直接 config 的对象，必须原样交出；这里只改外观。
        ``badge_text``：标题右侧的浅色小徽章（照视觉稿「可选」），留空则不显示。
        """
        box = tk.Frame(parent, bg=PANEL, highlightthickness=1, highlightbackground=LINE)
        box.pack(fill="x", pady=(0, SPACING["sm"]))
        row = tk.Frame(box, bg=PANEL)
        row.pack(fill="x", padx=SPACING["md"], pady=(6, 2))
        tk.Label(row, text=icon, bg=PANEL, fg=CINNABAR, font=F_ICON, width=2, height=1,
                 highlightthickness=1,
                 highlightbackground=CINNABAR).pack(side="left",
                                                    padx=(0, SPACING["sm"]))
        tk.Label(row, text=title, bg=PANEL, fg=INK,
                 font=F_SUBTITLE).pack(side="left")
        if badge_text:
            widgets.Badge(row, kind="trial", edition=self.edition,
                          text=badge_text).pack(side="left", padx=(SPACING["sm"], 0))
        # 选择按钮走公共零件（次级样式），不再用 ttk Ghost
        btn_frame = tk.Frame(row, bg=PANEL)
        btn_frame.pack(side="right")
        if btn2_text and cmd2:
            widgets.RoundButton(btn_frame, text=btn2_text, edition=self.edition,
                                style="secondary", height=SPACING["field_h"],
                                font=F_BTN,
                                command=cmd2).pack(side="left", padx=(0, SPACING["sm"]))
        widgets.RoundButton(btn_frame, text=btn_text, edition=self.edition,
                            style="secondary", height=SPACING["field_h"],
                            font=F_BTN,
                            command=cmd).pack(side="left")
        # 次行：选择状态 + 文件名 / 说明（长文件名按 280px 换行，不会撑破卡片）
        line2 = tk.Frame(box, bg=PANEL)
        line2.pack(fill="x", padx=SPACING["md"], pady=(0, 6))
        dot = tk.Label(line2, text="○", bg=PANEL, fg=MUTED, font=F_FOOT)
        dot.pack(side="left", padx=(0, SPACING["xs"]))
        lbl = tk.Label(line2, text="", bg=PANEL, fg=MUTED, font=F_FOOT)
        lbl.pack(side="left", padx=(0, SPACING["sm"]))
        name_lbl = tk.Label(line2, text=desc, bg=PANEL, fg=MUTED, font=F_FOOT,
                            anchor="w", justify="left", wraplength=420)
        name_lbl.pack(side="left", fill="x", expand=True)
        return box, name_lbl, dot, lbl

    # --------------------------------------------------------- right panel
    def _build_right(self, parent):
        """右卡：处理步骤（步进器 + 交付方式 + 如何工作 + 主操作 + 进度）。

        文案照 COPY_TABLE.md 第 3 节：标题「处理步骤」、右上「三步完成 / 论文格式检查与修正」、
        页码「1 / 3」、三步标题与说明、主按钮「下一步：开始检查 →」、底部信任卡。
        """
        hdr = tk.Frame(parent, bg=PANEL)
        hdr.pack(fill="x", pady=(0, 2))
        tk.Frame(hdr, bg=ACCENT, width=4, height=14).pack(side="left",
                                                          padx=(0, SPACING["sm"]))
        tk.Label(hdr, text="处理步骤", bg=PANEL, fg=INK, font=F_CARD_HDR).pack(side="left")
        self._step_counter = tk.Label(hdr, text="1 / 3", bg=PANEL, fg=MUTED, font=F_FOOT)
        self._step_counter.pack(side="right")
        # 右上斜排小字（视觉稿）：两行右对齐
        tk.Label(hdr, text="三步完成\n论文格式检查与修正", bg=PANEL, fg=MUTED,
                 font=F_FOOT, justify="right").pack(side="right", padx=(0, 14))

        # 步进器（装修包零件）：三个圆点 + 连接线 + 短标签；零件已暴露 circles/titles/lines，
        # 既有 _refresh_wizard / _on_step_done 的 config 调用原样可用。
        self._stepper = widgets.Stepper(
            parent, [(label, self.step_descs[i])
                     for i, (_mode, label) in enumerate(self.step_defs)],
            edition=self.edition)
        self._stepper.pack(fill="x", pady=(SPACING["sm"], SPACING["md"]))
        self._step_circle = self._stepper.circles
        self._step_title = self._stepper.titles
        self._step_line = self._stepper.lines

        # 交付方式（三选一：只批注 / 一键修正 / ①+② 都要）
        _mode_hint = tk.Label(parent, text="交付方式：① 只批注不改正文　② 直接改好　③ 两者都要",
                              bg=PANEL, fg=_ADV.primary, font=F_SMALL_B)
        _mode_hint.pack(anchor="w", fill="x")
        # 卡片一窄就按可用宽度折行（不设死 wraplength，宽窗仍保持单行不裁字）
        parent.bind("<Configure>",
                    lambda e: _mode_hint.configure(wraplength=max(200, e.width - 4)),
                    add="+")
        mode_row = tk.Frame(parent, bg=PANEL)
        mode_row.pack(fill="x", pady=(SPACING["xs"], SPACING["xs"]))
        self._btn_annotate = widgets.RoundButton(
            mode_row, text="① 只批注", edition=self.edition, style="primary",
            height=SPACING["field_h"],
            font=F_SMALL, command=lambda: self._set_fix_mode("annotate"))
        self._btn_annotate.pack(side="left", fill="x", expand=True,
                                padx=(0, SPACING["sm"]))
        self._btn_fix = widgets.RoundButton(
            mode_row, text="② 一键修正", edition=self.edition, style="secondary",
            height=SPACING["field_h"],
            font=F_SMALL, command=lambda: self._set_fix_mode("fix"))
        self._btn_fix.pack(side="left", fill="x", expand=True, padx=(0, SPACING["sm"]))
        self._btn_both = widgets.RoundButton(
            mode_row, text="③ ①+② 都要", edition=self.edition, style="secondary",
            height=SPACING["field_h"],
            font=F_SMALL, command=lambda: self._set_fix_mode("both"))
        self._btn_both.pack(side="left", fill="x", expand=True)
        tk.Frame(parent, bg=PANEL, height=SPACING["md"]).pack(fill="x")

        # 如何工作（装修包 info-panel 零件）
        widgets.InfoPanel(
            parent, edition=self.edition, title="如何工作",
            line1="上传论文 → 检查格式 → 查看问题 → 一键修正",
            line2="文件默认在本机处理，不会上传到服务器。").pack(fill="x")

        # 主操作：上一步 / 下一步（一屏一个主 CTA）
        btn_row = tk.Frame(parent, bg=PANEL)
        btn_row.pack(fill="x", pady=(SPACING["md"], 0))
        self._prev_btn = widgets.RoundButton(
            btn_row, text="上一步", edition=self.edition, style="secondary",
            height=SPACING["btn_h"],
            font=F_BTN, command=self._go_prev)
        self._prev_btn.pack(side="left", fill="x", expand=True, padx=(0, SPACING["sm"]))
        self._next_btn = widgets.RoundButton(
            btn_row, text="下一步：开始检查 →", edition=self.edition, style="primary",
            height=SPACING["btn_h"],
            font=F_BTN, command=self._run_step)
        self._next_btn.pack(side="left", fill="x", expand=True)

        # 底部信任卡（视觉稿）：复用同一 info-panel 零件，不另造第二套样式
        widgets.InfoPanel(
            parent, edition=self.edition, title="学术规范 · 专业高效",
            line1="精准识别格式问题，助力您的论文顺利通过审核。").pack(
                fill="x", pady=(SPACING["md"], 0))

        # 进度条：只在处理中由 _set_running(True) 显示，结束即收回，不留灰块
        self._progress_frame = tk.Frame(parent, bg=PANEL)
        self.progress = ttk.Progressbar(self._progress_frame, mode="indeterminate",
                                        length=300, style="Paper.Horizontal.TProgressbar")

        self._refresh_wizard()

    def _refresh_wizard(self):
        """根据 step_index / _errored / 当前交付方式的 _fix_phase 重绘时间线、计数与按钮三态。"""
        n = len(self.step_defs)
        # 第③步当前交付方式的状态（v1.0.15 起按模式独立：_fix_phase/_exported）
        _cur_phase = self._fix_phase.get(self._fix_mode, "idle")
        _cur_done = self._fix_mode in self._exported or _cur_phase in ("fixed", "saved")
        for i, (mode, label) in enumerate(self.step_defs):
            circ = self._step_circle[i]
            title = self._step_title[i]
            line = self._step_line[i]
            # 第三步（最后一步）当前交付方式已生成/已导出也视为完成态
            is_done = i < self.step_index or (
                i == n - 1 and _cur_done)
            if is_done:
                circ.config(bg=PANEL, fg=OKC, highlightbackground=OKC, text="✔")
                title.config(fg=INK)
                if line: line.config(bg=OKC)
            elif i == self.step_index and self._errored:
                circ.config(bg=PANEL, fg=ERRC, highlightbackground=ERRC, text="✕")
                title.config(fg=INK)
                if line: line.config(bg=_ADV.stepper_line)
            elif i == self.step_index:
                # 当前步：实心主色圆 + 白字（照装修包 stepper.svg 的选中态）
                circ.config(bg=ACCENT, fg="#FFFFFF", highlightbackground=ACCENT,
                            text=str(i + 1))
                title.config(fg=INK)
                if line: line.config(bg=_ADV.stepper_line)
            else:
                circ.config(bg=PANEL, fg=_ADV.muted, highlightbackground=_ADV.border,
                            text=str(i + 1))
                title.config(fg=MUTED)
                if line: line.config(bg=_ADV.stepper_line)
        self._step_counter.config(text="%d / %d" % (min(self.step_index + 1, n), n))
        if self.step_index >= n:
            self._next_btn.config(text="再处理一篇", command=self._reset_wizard, state="normal")
            self._prev_btn.config(text="上一步", state="disabled", command=self._go_prev)
        elif self.step_index == n - 1:
            # 第三步：依交付模式（批注副本 / 一键修正）显示按钮（由该模式的 _fix_phase 驱动）
            _fix_label = {"annotate": "生成批注副本",
                          "fix": "一键修正",
                          "both": "生成批注副本 + 已修正版"}[self._fix_mode]
            _phase = self._fix_phase.get(self._fix_mode, "idle")
            if self._fix_mode in self._exported:
                # v1.0.15：该交付方式已导出过 → 按钮显示「已完成」，点击提示已导出过，
                # 确认后可重新生成覆盖；不再静默可点、也不一刀切置灰（用户可随时切其他方式）。
                self._next_btn.config(text="已完成", command=self._re_export_prompt,
                                      state="disabled" if self.running else "normal")
                self._prev_btn.config(text="上一步",
                                      state="disabled" if self.running else "normal",
                                      command=self._go_prev)
            elif _phase == "idle":
                self._next_btn.config(text=_fix_label, command=self._run_step,
                                      state="disabled" if self.running else "normal")
                self._prev_btn.config(text="上一步",
                                      state="disabled" if (self.running or self.step_index == 0) else "normal",
                                      command=self._go_prev)
            elif _phase == "fixed":
                self._next_btn.config(text="保存结果", command=self._export_fix,
                                      state="disabled" if self.running else "normal")
                self._prev_btn.config(text="上一步", state="normal", command=self._go_prev)
            else:  # saved（防御分支：正常已归入 _exported）
                self._next_btn.config(text="完成", state="disabled")
                self._prev_btn.config(text="再处理一篇", state="normal",
                                      command=self._reset_wizard)
        else:
            self._next_btn.config(text="下一步：开始检查 →", command=self._run_step,
                                  state="disabled" if self.running else "normal")
            self._prev_btn.config(text="上一步",
                                  state="disabled" if (self.running or self.step_index == 0) else "normal",
                                  command=self._go_prev)

    def _set_bar(self, state, hint=None):
        """底部状态栏：idle / running / done / error。"""
        cmap = {"idle": OKC, "running": RUN, "done": OKC, "error": ERRC}
        # 左侧固定身份行（状态由圆点颜色 + 右侧状态文字表达）；右侧提示空闲时留空，
        # 免得「请按步骤操作」在一条状态栏里出现两遍。
        _ident = "论文格式医生 · 导师版 v%s · 本机处理" % APP_VERSION
        tmap = {
            "idle": (_ident, ""),
            "running": (_ident, hint or widgets.COPY["processing"]),
            "done": (_ident, widgets.COPY["done"]),
            "error": (_ident, widgets.COPY["error"]),
        }
        left, right = tmap.get(state, tmap["idle"])
        c = cmap.get(state, MUTED)
        self.bar_left.config(text=left, fg=c)
        self.bar_dot.config(fg=c)
        self.bar_right.config(text=right)

    def _on_step_done(self, idx, mode):
        self._step_circle[idx].config(bg=PANEL, fg=OKC, highlightbackground=OKC, text="✔")
        self._step_title[idx].config(fg=INK)
        if self._step_line[idx]:
            self._step_line[idx].config(bg=OKC)
        if mode == "fix":
            # 第三步：修正完成→进入“保存修正后论文”子状态（先修正、后导出）。
            # 不前进到“再处理一篇”，按钮由该交付方式的 _fix_phase 驱动为「保存修正后论文」。
            self._fix_phase[self._fix_mode] = "fixed"
            self._set_status("修正完成，请点击「保存修正后论文」", OKC)
            self._set_bar("done")
        else:
            self.step_index = idx + 1
            self._set_status(widgets.COPY["done"], OKC)
            self._set_bar("done")
        self._refresh_wizard()

    def _on_step_error(self, idx, mode, err):
        self._errored = True
        if mode == "fix":
            self._fix_phase[self._fix_mode] = "idle"
        self._set_status(widgets.COPY["error"], ERRC)
        self._set_bar("error")
        self._refresh_wizard()
        title = "处理出错"
        msg = err
        if mode == "fix":
            title = "一键修正失败"
            log_path = os.path.join(tempfile.gettempdir(), "tfd_fix_error.log")
            msg = ("一键修正未能完成，论文原文件未被改动。\n\n"
                   "错误信息：\n%s\n\n"
                   "完整报错已记录到：\n%s\n\n"
                   "请把这段信息与该日志文件发给客服，以便定位原因。"
                   % (err, log_path)) + HELP_HINT
        messagebox.showerror(title, msg)

    # --------------------------------------------------------------- picks
    def _update_thesis_status(self):
        """根据已选论文数刷新『批量导入论文』那行的显示：已选择 N 篇论文。"""
        n = len(getattr(self, "thesis_paths", []))
        if n == 0:
            self._thesis_lbl.config(text="未选择论文", fg=MUTED)
            self._thesis_name.config(
                text="可选择文件夹，也可选择多个文件", fg=MUTED)
            self._thesis_dot.config(text="○", fg=MUTED)
            return
        self._thesis_dot.config(text="✓", fg=OKC)
        if n == 1:
            self._thesis_lbl.config(text="已选择 1 篇论文", fg=INK)
            self._thesis_name.config(
                text=self._clip(os.path.basename(self.thesis_paths[0])), fg=INK)
        else:
            self._thesis_lbl.config(text="已选择 %d 篇论文" % n, fg=INK)
            self._thesis_name.config(
                text=self._clip("已选 %d 篇：%s …" % (n, os.path.basename(self.thesis_paths[0]))), fg=INK)

    def _pick_input(self):
        paths = filedialog.askopenfilenames(
            title="选择待导入论文（可多选批量导入）",
            filetypes=[("Word 文档", "*.docx *.doc *.wps"), ("所有文件", "*.*")])
        if paths:
            self.thesis_paths = list(paths)
            # 兼容旧引用：thesis_path 取首篇，供报告命名等使用
            self.thesis_path.set(self.thesis_paths[0])
            self._update_thesis_status()

    def _pick_folder(self):
        """选择文件夹：自动收齐其中所有 Word 文档（.docx/.doc/.wps）。"""
        folder = filedialog.askdirectory(
            title="选择论文所在文件夹（自动导入其中所有 Word 文档）")
        if not folder:
            return
        exts = (".docx", ".doc", ".wps")
        files = []
        for fn in sorted(os.listdir(folder)):
            if fn.lower().endswith(exts) and os.path.isfile(os.path.join(folder, fn)):
                files.append(os.path.join(folder, fn))
        if not files:
            messagebox.showinfo(
                "文件夹内无文档",
                "该文件夹中没有找到 Word 文档（.docx / .doc / .wps）。\n"
                "请确认论文已放在此文件夹下。")
            return
        self.thesis_paths = files
        self.thesis_path.set(self.thesis_paths[0])
        self._update_thesis_status()

    def _pick_template(self):
        p = filedialog.askopenfilename(
            title="选择学校模板",
            filetypes=[("Word 文档", "*.docx *.doc *.wps"), ("所有文件", "*.*")])
        if p:
            self.template_path.set(p)
            self._template_name.config(text=self._clip(os.path.basename(p)), fg=INK)
            self._tpl_dot.config(text="✓", fg=OKC)
            self._tpl_lbl.config(text="模板已选择", fg=INK)
            # 换了新模板：重新允许提取画像（去掉"放弃"标记）
            self._profile_abandoned = False
            self._profile_confirmed = False
            if self._tpl_locked:
                self._save_template_lock(p)

    # --------------------------------------------------- 模板锁定（记住模板）
    def _template_lock_path(self):
        return os.path.join(os.path.expanduser("~"), ".tfd_mentor_license", "template_lock.json")

    def _save_template_lock(self, path):
        try:
            p = self._template_lock_path()
            os.makedirs(os.path.dirname(p), exist_ok=True)
            with open(p, "w", encoding="utf-8") as f:
                json.dump({"path": path}, f, ensure_ascii=False)
            self._tpl_locked = True
            self._paint_lock_btn()
            # 提示条文案照 COPY_TABLE 第 5 节
            self._flash_lock(widgets.COPY["toast_saved"])
        except Exception as e:
            self._debug("[模板锁定保存失败] " + str(e))

    def _load_template_lock(self):
        try:
            p = self._template_lock_path()
            if not os.path.isfile(p):
                return None
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f).get("path")
        except Exception:
            return None

    def _delete_template_lock(self):
        try:
            p = self._template_lock_path()
            if os.path.isfile(p):
                os.remove(p)
        except Exception:
            pass
        self._tpl_locked = False
        self._paint_lock_btn()
        self._flash_lock("已取消记住")

    def _paint_lock_btn(self):
        """根据 _tpl_locked 刷新『记住此模板』按钮：实底=已记住，描边=未记住。

        旧的橙底/绿底 tk.Button 是上一版遗留的警示配色，改用公共零件后与全站按钮同款。
        """
        if self._tpl_locked:
            self._lock_btn.set_text("✓ 已记住（点此取消）")
            self._lock_btn.set_style("primary")
        else:
            self._lock_btn.set_text("🔒 记住此模板")
            self._lock_btn.set_style("secondary")

    def _flash_lock(self, msg):
        """在按钮旁短暂显示提示文字，1.8 秒后清空。"""
        self._lock_tip.config(text=msg)
        self._lock_tip.after(1800, lambda: self._lock_tip.config(text=""))

    def _on_lock_click(self):
        """点『记住此模板』按钮：未记住→保存锁定；已记住→取消。"""
        if self._tpl_locked:
            self._delete_template_lock()
            return
        p = self.template_path.get().strip()
        if not (p and os.path.isfile(p)):
            messagebox.showinfo("先选择模板",
                                "请先选择学校模板，再点“记住此模板”。")
            return
        self._save_template_lock(p)

    def _restore_template_lock(self):
        """启动时：若已记住模板且文件仍在，自动载入并后台预提取格式画像。"""
        p = self._load_template_lock()
        if not p or not os.path.isfile(p):
            if p:
                self._delete_template_lock()  # 模板被移动/删除，清理失效锁定
            return
        self._tpl_locked = True
        self._paint_lock_btn()
        self.template_path.set(p)
        self._template_name.config(text=self._clip(os.path.basename(p)), fg=INK)
        self._tpl_dot.config(text="✓", fg=OKC)
        self._tpl_lbl.config(text="模板已选择（已记住）", fg=INK)
        self._flash_lock("已自动载入记住的模板")
        threading.Thread(target=self._extract_profile, args=(p,), daemon=True).start()

    def _clear_profile(self):
        self.profile_path.set("")
        self._update_profile_box()

    def _paint_mode_buttons(self):
        """按当前 _fix_mode 高亮三个交付方式按钮（选中=主色实底，未选=描边）。

        样式走公共零件 RoundButton.set_style，不再给旧 tk.Button 逐色 config
        （旧写法既写死了第二套色，也做不出新设计的描边/实底两态）。
        """
        if getattr(self, "_btn_annotate", None) is None:
            return
        for mode, btn in (("annotate", self._btn_annotate),
                          ("fix", self._btn_fix),
                          ("both", self._btn_both)):
            btn.set_style("primary" if self._fix_mode == mode else "secondary")

    def _set_fix_mode(self, mode):
        """切换第③步交付方式：annotate=只批注不修改（默认）/ fix=一键修正 / both=①+② 都要。"""
        self._fix_mode = mode
        self._fix_mode_var.set(mode)
        self._paint_mode_buttons()
        self._refresh_wizard()

    def _re_export_prompt(self):
        """已导出的交付方式：点击「已完成」时提示已导出过，确认后可重新生成导出（覆盖）。

        v1.0.15：三个交付方式可随意切换、各自独立导出；已导出的方式不再静默重复导出，
        而是先提示已保存位置，由客户决定是否重新生成覆盖。
        """
        mode = self._fix_mode
        saved = (self._exported or {}).get(mode) or set()
        label = {"annotate": "① 只批注·不改原稿",
                 "fix": "② 一键修正·直接改好",
                 "both": "③ ①+② 都要"}[mode]
        files = "\n".join(os.path.basename(p) for p in sorted(saved)) or "（记录缺失）"
        if not messagebox.askyesno(
                "已导出过",
                "「%s」已导出过，结果已保存：\n%s\n\n"
                "确定要重新生成并导出（覆盖）吗？\n"
                "选「否」则可直接切换其他交付方式继续。" % (label, files)):
            return
        # 重新生成：退回待生成状态并清空旧导出记录（生成→保存后重新记录），
        # 走完整第③步流程（产出重新写入 _fix_outs 后再导出）
        self._fix_phase[mode] = "idle"
        self._exported.pop(mode, None)
        self._refresh_wizard()
        self._run_step()

    # ---------------------------------------------------------------- run（向导）
    def _go_prev(self):
        """上一步：回退一个步骤，该步及其后的进度重置为待办，可重新执行。

        第三步（修正）内的子状态：先退回“未修正”，停留在第三步便于重新修正，
        而非跳回检查步骤。
        """
        if self.running or self.step_index <= 0:
            return
        if self.step_index == len(self.step_defs) - 1 and self._fix_phase.get(self._fix_mode, "idle") != "idle":
            # 上一步 = 明确重做：退回该交付方式的"待生成"状态，并清空其导出记录，
            # 按钮恢复为「生成…」而非「已完成」（否则 exported 残留会让按钮卡在已完成）。
            self._fix_phase[self._fix_mode] = "idle"
            self._exported.pop(self._fix_mode, None)
            self._errored = False
            self._set_status("请按步骤操作", MUTED)
            self._set_bar("idle")
            self._refresh_wizard()
            return
        self.step_index -= 1
        self._errored = False
        # v1.3.49：回退后允许重新确认画像（去掉"放弃"标记）
        self._profile_abandoned = False
        self._set_status("请按步骤操作", MUTED)
        self._set_bar("idle")
        self._refresh_wizard()

    def _run_step(self):
        if self._dialog_open:
            return  # 保存对话框已打开，防连点重复弹窗
        if self.running:
            messagebox.showinfo("正在处理", "上一步还在处理中，请稍候…")
            return
        if self.step_index >= len(self.step_defs):
            self._reset_wizard()
            return
        idx = self.step_index
        mode = self.step_defs[idx][0]

        # 校验：论文文件必选（可批量，模板可选）
        if not self.thesis_paths:
            messagebox.showerror("缺少输入", "请先选择“批量导入论文”（可多选批量导入）。")
            return

        # 批量（≥2 篇）：检查 / 修正都走批量。先让客户选输出文件夹，再线程批量处理（避免逐个弹保存框）。
        # v1.0.14：修正此前"仅 fix 模式批量、check 模式选多篇只处理首篇"导致只输出一篇报告的 bug。
        if len(self.thesis_paths) > 1:
            out_dir = filedialog.askdirectory(
                title="选择批量输出文件夹（%s报告将保存在此）"
                       % ("检查" if mode == "check" else "处理结果"))
            if not out_dir:
                return
            self._errored = False
            self.running = True
            self._set_running(True)
            self._set_status("正在批量处理 %d 篇论文…" % len(self.thesis_paths), RUN)
            self._set_bar("running", "批量处理")
            threading.Thread(target=self._batch_worker, args=(idx, mode, out_dir),
                             daemon=True).start()
            return

        # 单篇路径（维持原有逻辑）；第③步「保存位置」放到修正【完成之后】再弹
        src = self.thesis_paths[0]
        dst = None

        self._errored = False
        self.running = True
        self._set_running(True)
        status = {"profile": "正在提取学校模板要求…",
                  "check": "正在检查论文格式…",
                  "fix": {"annotate": "正在生成批注副本…",
                          "fix": "正在按学校要求修正论文…",
                          "both": "正在生成批注副本 + 已修正版…"}[self._fix_mode]}[mode]
        self._set_status(status, RUN)
        self._set_bar("running", self.step_defs[idx][1])
        # 点亮当前步
        self._step_circle[idx].config(bg="#ffffff", fg=ACCENT,
                                      highlightbackground=ACCENT, text=str(idx + 1))
        self._step_title[idx].config(fg=INK)
        threading.Thread(target=self._worker, args=(idx, mode, src, dst), daemon=True).start()

    def _trial_ok(self):
        """试用门禁：已激活直接放行；未激活则校验次数/弹确认，扣减失败拦截。
        在 worker 线程调用（内部弹窗走主线程 Event 同步）。返回 True=可继续。"""
        licensed = trial.is_licensed()
        if licensed:
            return True
        left = trial.trials_left()
        if left <= 0:
            self.root.after(0, self._show_trial_exhausted)
            return False
        if not self._ask_trial_confirm(left):
            return False
        if not trial.consume_trial():
            self.root.after(0, self._show_trial_exhausted)
            return False
        return True

    def _worker(self, idx, mode, src, dst=None):
        try:
            if mode == "profile":
                self._do_profile(src)
                self.root.after(0, lambda: self._on_step_done(idx, mode))
            else:
                docx_path, note = engine.normalize_input(src)
                if note:
                    self._debug(note)
                if mode == "check":
                    # 第二步：先在「学校模板要求」确认窗核对/修改（修改写回画像），
                    # 确认后才开始检查；保存报告放到检查完成后由主线程弹框（_save_check_report）。
                    confirmed = self._confirm_profile_if_needed()
                    if confirmed is None:
                        self.profile_path.set("")   # 客户放弃使用画像 → 按通用规范检查
                    report = self._do_check(src, docx_path)
                    self.root.after(0, lambda: self._save_check_report(report, idx))
                elif mode == "fix":
                    # 修正前同样先确认/修改画像（第一次进修正时）
                    confirmed = self._confirm_profile_if_needed()
                    if confirmed is None:
                        self.profile_path.set("")
                    # 试用门禁只校验一次（两种都要时也只扣一次）
                    if not self._trial_ok():
                        self.root.after(0, self._on_trial_blocked)
                        return
                    self._do_fix(src, docx_path, dst)
                    self.root.after(0, lambda: self._on_step_done(idx, mode))
        except Exception as e:
            self._debug("[错误] " + str(e))
            self.root.after(0, lambda: self._on_step_error(idx, mode, str(e)))
        finally:
            self.running = False
            # 步骤处理结束：清理旧格式转换产生的临时目录（客户无感，不残留 tfd_conv_*）
            engine.cleanup_conv_dirs()
            self.root.after(0, lambda: self._set_running(False))

    def _set_running(self, running):
        def _apply():
            if running:
                # v1.0.14：进度条显示到"上一步/下一步"按钮下方的独立容器。
                self._progress_frame.pack(fill="x", pady=(10, 0))
                self.progress.pack(fill="x", padx=2, pady=2)
                self.progress.start(12)
                self._next_btn.config(state="disabled")
            else:
                self.progress.stop()
                self.progress.pack_forget()
                self._progress_frame.pack_forget()
                self._batch_hide()
                # v1.0.14：结束态用 _refresh_wizard 重新派生按钮三态——
                # 修复此前无条件 state="normal" 把"完成"态（应置灰）错误恢复成可点的 bug。
                self._refresh_wizard()
        self.root.after(0, _apply)

    # ------------------------------------------------- 批量处理进度（纯 UI 刷新）
    def _batch_tick(self, done, total, name):
        """批量进度块：只刷新渲染对象，不参与也不改变任何处理逻辑。

        调用点在 worker 线程 → 一律经 after(0) 切回主线程再碰 Tk（与 _set_status 同规矩）。
        """
        self.root.after(0, lambda: self._batch_tick_ui(done, total, name))

    def _batch_tick_ui(self, done, total, name):
        prog = getattr(self, "_batch_prog", None)
        if prog is None:
            return
        if not prog.winfo_ismapped():
            prog.pack(fill="x", pady=(0, 8), before=self._thesis_box)
        self._batch_count.config(text="%d / %d" % (done, total))
        self._batch_file.config(text="正在处理%s" % name)
        try:
            self._batch_bar.config(value=(100.0 * done / total) if total else 0)
        except Exception:
            pass

    def _batch_hide(self):
        """收起批量进度块（空闲时不留空框）。"""
        prog = getattr(self, "_batch_prog", None)
        if prog is not None and prog.winfo_ismapped():
            prog.pack_forget()

    # --------------------------------------------------- profile 提取与确认
    def _extract_profile(self, src):
        """提取学校模板要求并弹确认页（客户可修改）。返回画像路径或 None（放弃）。"""
        t_docx, note = engine.normalize_input(src)
        if note:
            self._debug(note)
        # 画像 json 放到系统临时目录（不污染客户源目录），按源文件 hash 命名避免多论文覆盖
        out = os.path.join(tempfile.gettempdir(),
                           "tfd_profile_%s.json"
                           % hashlib.sha1(src.encode("utf-8")).hexdigest()[:10])
        text = engine.run_build_profile(t_docx, out)
        try:
            profile = json.loads(text)
        except Exception:
            profile = None
        if profile is None:
            self.profile_path.set("")
            return None
        # 提取成功先不弹确认页：确认页挪到「检查/修正」步骤执行前统一弹出，
        # 客户在真正处理前核对/修改（修改会写回画像并生效）。
        self.profile_path.set(out)
        self._profile_confirmed = False
        self._profile_abandoned = False   # 新画像就绪：清除"放弃"标记
        self.root.after(0, self._update_profile_box)
        return out

    def _ensure_profile_ready(self):
        # v1.3.49：客户已明确"放弃"画像 → 本次会话不再自动重新提取（按通用规范处理），
        # 否则放弃后又会从模板重新提取，放弃形同虚设。
        if self._profile_abandoned:
            return None
        p = self.profile_path.get().strip()
        if p and os.path.isfile(p):
            return p
        t = self.template_path.get().strip()
        if t and os.path.isfile(t):
            return self._extract_profile(t)
        return None

    def _confirm_profile_if_needed(self):
        """检查/修正前：若画像存在且尚未确认，弹「学校模板要求」确认窗（可修改）。

        客户修改的字段会写回画像 json，后续检查/修正均按修改后的要求执行。
        返回画像路径；客户选择“放弃”时返回 None（本次按通用规范处理）。
        """
        p = self._ensure_profile_ready()
        if not p or not os.path.isfile(p) or self._profile_confirmed:
            return p
        try:
            with open(p, encoding="utf-8") as f:
                profile = json.load(f)
        except Exception:
            return p
        ok, edits = self._ask_profile_confirm(profile)
        if not ok:
            # v1.3.49：客户点"放弃"→ 本次会话不再自动重新提取画像，按通用规范处理
            self._profile_abandoned = True
            return None
        self._profile_abandoned = False
        if edits:
            for path_keys, value in edits:
                _deep_set(profile, path_keys, value)
                twips_key = _MARGIN_TWIPS.get(tuple(path_keys))
                if twips_key:
                    try:
                        _deep_set(profile, ("spec", "page", twips_key),
                                  str(int(float(value) * _TWIPS_PER_CM)))
                    except (ValueError, TypeError):
                        pass
            try:
                with open(p, "w", encoding="utf-8") as f:
                    json.dump(profile, f, ensure_ascii=False, indent=2)
            except Exception:
                pass
        self._profile_confirmed = True
        return p

    def _ask_profile_confirm(self, profile):
        ev = threading.Event()
        box = {}

        def show():
            try:
                ok, edits = self._profile_confirm_dialog(profile)
                box["ok"] = ok
                box["edits"] = edits
            except Exception:
                box["ok"] = False
                box["edits"] = []
            finally:
                ev.set()

        self.root.after(0, show)
        # 注意：不能用 ev.wait(timeout=...) 限时等待——客户读完弹窗再点“确认”必然
        # 超过限定时长，导致 ev 超时、box 仍为空、确认结果被丢弃（画像被清空、退回通用规范）。
        # 弹窗通过 wait_window 阻塞主线程，客户关闭后 show() 才返回并 ev.set()，
        # 这里无限等待直到客户做出选择即可（工作线程等待、主线程照常处理弹窗，无死锁）。
        ev.wait()
        return box.get("ok", False), box.get("edits", [])

    def _profile_confirm_dialog(self, profile):
        """“学校模板要求 · 请确认”弹窗：只读摘要 + 可修改关键项。返回 (ok, edits)。"""
        result = {"ok": False, "edits": []}
        t = self._theme
        # 高度按屏幕夹取，避免 768 高屏把「确认 / 放弃」按钮挤出屏外（overrideredirect 无标题栏拖不动）
        try:
            _sh = self.root.winfo_screenheight()
        except Exception:
            _sh = 768
        _dlg_h = min(780, max(520, _sh - 100))
        top = ModalShell(self.root, edition=self.edition,
                         title="学校模板要求 · 请确认", width=680, height=_dlg_h)
        tk.Label(top.body, text="已提取出学校模板的格式要求", bg=t.modal_fill, fg=t.modal_title,
                 font=t.serif(TYPE["section_title"], bold=True)).pack(anchor="w", pady=(8, 2))
        tk.Label(top.body, text="请核对是否与学校规定一致；如有不准，可直接修改后确认。",
                 bg=t.modal_fill, fg=t.modal_sub,
                 font=t.sans(TYPE["caption"])).pack(anchor="w", pady=(0, 6))

        # v1.3.46：批注来源提示——有批注=绿色"以批注为准"；无批注=红色警告（防误解、减纠纷）
        _n_cmt = int((profile or {}).get("comment_count", 0) or 0)
        if _n_cmt > 0:
            src_note = tk.Frame(top.body, bg=t.success_fill, highlightthickness=1,
                                highlightbackground=t.success_border)
            src_txt = "已读取学校模板批注 %d 条 —— 以下要求以【批注】为准（最权威）。" % _n_cmt
            src_fg = t.success_text
        else:
            src_note = tk.Frame(top.body, bg=t.error_fill, highlightthickness=1,
                                highlightbackground=t.error_border)
            src_txt = ("该模板【未检测到批注】。学校批注是最权威的格式要求；"
                       "无批注时以下要求来自模板样式定义 / 通用规范，"
                       "可能与学校规定有出入，请仔细核对后再确认。")
            src_fg = t.error_text
        src_note.pack(fill="x", pady=(0, 6))
        tk.Label(src_note, text=src_txt, bg=src_note.cget("bg"), fg=src_fg,
                 font=t.sans(TYPE["caption"]), justify="left", anchor="w",
                 wraplength=560).pack(fill="x", padx=10, pady=6)

        sum_f = tk.Frame(top.body, bg=t.surface, highlightthickness=1, highlightbackground=t.border)
        sum_f.pack(fill="x", pady=3)
        sum_txt = tk.Text(sum_f, height=6, wrap="word", bg=t.surface, fg=t.modal_sub,
                          font=t.sans(TYPE["caption"]), relief="flat", padx=10, pady=6)
        sum_txt.insert("1.0", "\n".join(_profile_summary(profile)))
        sum_txt.config(state="disabled")
        sum_txt.pack(fill="x")

        tk.Label(top.body, text="如需修正，直接修改下列项目（留空表示保持提取结果）",
                 bg=t.modal_fill, fg=t.primary, font=t.sans(TYPE["body"], bold=True)).pack(
            anchor="w", pady=(8, 2))

        # 内容区（含可滚动表单）占满剩余空间，按钮固定底部
        content = tk.Frame(top.body, bg=t.modal_fill)
        content.pack(fill="both", expand=True, pady=3)
        # 表单区：canvas 与滚动条同在一个 frame 内，滚动条贴右侧整个高度
        form_area = tk.Frame(content, bg=t.modal_fill)
        form_area.pack(fill="both", expand=True)
        canvas = tk.Canvas(form_area, bg=t.modal_fill, highlightthickness=0)
        vbar = ttk.Scrollbar(form_area, orient="vertical", command=canvas.yview)
        form = tk.Frame(canvas, bg=t.modal_fill)
        form.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=form, anchor="nw")
        canvas.configure(yscrollcommand=vbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        vbar.pack(side="right", fill="y")

        # v1.3.76：确认页表单滚动区同样支持滚轮（15 个字段，鼠标翻页更顺手）
        def _form_wheel(e):
            d = getattr(e, "delta", 0) or 0
            step = -3 if d > 0 else 3
            if abs(d) >= 120:
                step = -int(d / 120) * 3
            canvas.yview_scroll(step, "units")
        for w in (canvas, form):
            w.bind("<MouseWheel>", _form_wheel)
            w.bind("<Button-4>", _form_wheel)
            w.bind("<Button-5>", _form_wheel)

        entries = {}
        for i, (label, candidates, kind) in enumerate(EDIT_FIELDS):
            tk.Label(form, text=label, bg=t.modal_fill, fg=t.modal_sub,
                     font=t.sans(TYPE["body"])).grid(
                row=i, column=0, sticky="e", padx=(0, 10), pady=3)
            cur = _field_value(profile, candidates)
            if kind == "align":
                var = tk.StringVar(value=ALIGN_DISPLAY.get(cur, cur))
                cb = ttk.Combobox(form, textvariable=var, width=20,
                                  font=t.sans(TYPE["body"]),
                                  values=list(ALIGN_DISPLAY.values()), state="readonly")
                cb.grid(row=i, column=1, sticky="w", pady=3)
                entries[i] = (candidates, var, "align", var.get())
            elif kind == "ref":
                has_ref = bool(cur)
                var = tk.StringVar(value="已提取" if has_ref else "未提取（按通用规范检查）")
                tk.Entry(form, textvariable=var, width=28, font=t.sans(TYPE["body"]),
                         relief="flat", bd=0, bg=t.modal_fill, state="disabled",
                         disabledforeground=t.success_text if has_ref else t.muted).grid(
                    row=i, column=1, sticky="w", pady=3)
                entries[i] = (candidates, var, "ref", var.get())
            else:
                var = tk.StringVar(value=cur)
                tk.Entry(form, textvariable=var, width=22, font=t.sans(TYPE["body"]),
                         relief="solid", bd=1, highlightthickness=1,
                         highlightbackground=t.select_border).grid(
                    row=i, column=1, sticky="w", pady=3)
                entries[i] = (candidates, var, "text", var.get())

        _NUM_FIELDS = ("indent_chars", "line_val", "top_cm", "bottom_cm",
                       "left_cm", "right_cm", "before_pt", "after_pt", "hanging_cm")

        def on_confirm():
            edits = []
            for (candidates, var, kind, init_text) in entries.values():
                if kind == "ref":
                    continue  # 参考文献格式为只读提示，不参与修改
                text = var.get().strip()
                # v1.3.43：只写回【真正被修改】的字段——此前把"所有非空字段"全部写回，
                # 导致未修改的 sz（显示为磅）被当成半磅写回、数值字段变字符串，
                # 修正引擎 %d 崩溃/字号错乱（用户实测"一键修正没改论文"的根因）。
                if text == init_text or not text:
                    continue
                if kind == "align":
                    code = ALIGN_CODE.get(text, text)
                    edits.append((candidates[0], code))
                    continue
                k = candidates[0][-1]
                if k == "sz":
                    # 确认页显示为磅（_field_value ÷2），写回需还原为半磅（×2）。
                    # v1.3.49：支持输入中文字号名（如"小四"→12磅）或数字磅值。
                    pt = _cn_size_to_pt(text)
                    if pt is None:
                        continue  # 非法输入：保持提取结果
                    edits.append((candidates[0], str(int(round(pt * 2)))))
                elif k in _NUM_FIELDS:
                    try:
                        float(text)
                    except (TypeError, ValueError):
                        continue  # 非法输入：保持提取结果
                    edits.append((candidates[0], text))
                    # v1.3.47：客户填了首行缩进 → 同时落定 indent_type="first"，
                    # 否则画像无 indent_type 时引擎不会套用缩进（用户实测踩坑）
                    if k == "indent_chars":
                        edits.append((candidates[0][:-1] + ("indent_type",), "first"))
                    # v1.3.48：客户填了行距(磅) → 同时落定 line_rule="exact"（固定值行距），
                    # 否则画像无 line_rule 时引擎不设行距（实测发现同款坑）
                    if k == "line_val":
                        edits.append((candidates[0][:-1] + ("line_rule",), "exact"))
                else:
                    edits.append((candidates[0], text))
            result["ok"] = True
            result["edits"] = edits
            top.destroy()

        def on_cancel():
            result["ok"] = False
            top.destroy()

        btns = tk.Frame(top.body, bg=t.modal_fill)
        btns.pack(side="bottom", fill="x", pady=(8, 2))
        widgets.RoundButton(btns, text="确认，使用此要求", edition=self.edition, style="primary",
                            height=SPACING["btn_h"], font=t.sans(TYPE["btn"], bold=True),
                            command=on_confirm).pack(side="left", padx=(0, 8))
        widgets.RoundButton(btns, text="放弃（不使用画像）", edition=self.edition, style="secondary",
                            height=SPACING["btn_h"],
                            font=t.sans(TYPE["btn_small"], bold=True),
                            command=on_cancel).pack(side="left")
        top.after(10, lambda: canvas.yview_moveto(0))
        top.center_on(self.root)
        top.wait_window()
        return result["ok"], result["edits"]

    # ------------------------------------------------------------ 各步骤
    def _do_profile(self, src):
        """第①步：提取【学校模板】的格式要求（不是从论文提取）。

        画像必须来自学校模板才有意义；若未选模板，则没有“学校要求”可提取，
        走通用规范并明确告知客户，且不弹一个空的“确认模板要求”框。
        """
        tpl = self.template_path.get().strip()
        if tpl and os.path.isfile(tpl):
            out = self._extract_profile(tpl)
            if out:
                self._debug("学校模板画像已保存：" + out)
            else:
                self._debug("学校模板要求提取失败（模板可能无样式/批注）")
        else:
            # 未选模板：无学校要求可提取，按通用规范处理，标记已确认避免后续弹空框
            self.profile_path.set("")
            self._profile_confirmed = True
            self.root.after(0, lambda: messagebox.showinfo(
                "无需提取学校要求",
                "未选择学校模板，将按通用论文格式规范检查 / 修正。\n\n"
                "如需按学校具体要求处理，请点“上一步”回到第①步，先选择学校模板再重做。"))
            self.root.after(0, self._update_profile_box)

    def _do_check(self, src, docx_path):
        profile = self._ensure_profile_ready()
        report = engine.run_check(docx_path, profile_path=profile)
        self._debug(report)
        return report

    def _save_check_report(self, report, idx):
        """主线程：检查完成后让客户选择保存位置，写入 Word 报告并收尾该步骤。

        同一会话内已保存过检查报告时，再次保存前弹“已保存过，是否再次保存”确认。
        """
        base = _base_no_ext(self.thesis_path.get().strip() or "report")
        dst = filedialog.asksaveasfilename(
            title="选择检查报告保存位置",
            initialfile=os.path.basename(base) + "_格式检查报告.docx",
            initialdir=os.path.dirname(base) or None,
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")])
        saved = False
        if dst:
            if self._check_report_saved:
                again = messagebox.askyesno(
                    "已保存过",
                    "检查报告之前已保存过：\n%s\n\n确定要再次保存（覆盖）吗？"
                    % os.path.basename(self._check_report_saved))
                if not again:
                    dst = None
            if dst:
                try:
                    engine.md_to_docx(report, dst)
                    self._check_report_saved = dst
                    saved = True
                except Exception as e:
                    messagebox.showerror("保存失败", str(e))
                    dst = None
        else:
            self._debug("客户未选择保存位置，检查报告未落盘")
        self._on_step_done(idx, "check")
        if saved:
            self._show_check_done(dst)

    def _do_fix(self, src, docx_path, dst):
        # 授权被后台撤销（退款锁死）：直接拦截，禁止继续修正
        if license.is_revoked():
            self.root.after(0, self._handle_revoked)
            return False
        # 第③步修正：按当前 _fix_mode 产出（both=批注副本+已修正版各一套，不重跑论文）。
        # 产出先落临时目录，待主线程 _export_fix 让客户选保存位置；结果存入 self._fix_outs。
        # 注意：试用门禁（_trial_ok）已由调用方在进此方法前统一校验一次，避免"两种都要"重复扣次数。
        licensed = trial.is_licensed()
        author = self.author_var.get().strip() or None
        modes = ["annotate", "fix"] if self._fix_mode == "both" else [self._fix_mode]
        outs = {}
        profile = self._ensure_profile_ready()  # 画像循环外只提取一次（①+② 都要时两套共用）
        for mode in modes:
            annotate = (mode == "annotate")
            suffix = "_批注副本" if annotate else "_已修正"
            if not dst:
                tmp = tempfile.mkdtemp(prefix="tfd_fix_")
                base = os.path.join(tmp, os.path.splitext(os.path.basename(src))[0])
                odst = base + suffix + ".docx"
                rep = base + suffix + "_修改报告.docx"
                chk = base + suffix + "_检查报告.docx"
            else:
                base_dst = _base_no_ext(dst)
                odst = base_dst + suffix + ".docx"
                rep = base_dst + suffix + "_修改报告.docx"
                chk = base_dst + suffix + "_检查报告.docx"
            xlsx = os.path.splitext(odst)[0] + "_修改明细.xlsx"  # 导师版偏好 Excel 明细
            # 主交付物：修正/批注后的论文 + 修改明细 Excel（run_* 内部已写盘）
            if annotate:
                # 生成批注副本：只标不改，comments 作者=署名
                report = engine.run_annotate(
                    docx_path, odst, profile_path=profile, author=author, xlsx_path=xlsx)
                rep = None  # 批注副本无"修改报告"（本身未改动）
                chk = None  # 批注副本也无"检查报告"，置空避免导出时空 copy 引发"导出失败"
            else:
                report = engine.run_fix_headings(
                    docx_path, odst, profile_path=profile,
                    report_docx=rep, add_comments=True, author=author, xlsx_path=xlsx)
            if not licensed:
                # 试用版：给产出文档加水印+只读保护，并让客户知道正式版可编辑无水印
                if watermark.apply_watermark(odst):
                    report += ("\n\n> 本预览版带水印且为【只读】文档（编辑需密码，仅作效果预览）；"
                               "激活码解锁后输出可编辑无水印正式版，可一键交稿。\n")
                else:
                    self._debug("[试用水印注入失败，已跳过]")
            self._debug(report)
            # 修改明细报告若因引擎内报告环节异常未落盘，置空（主交付物不受影响）
            if rep and not os.path.isfile(rep):
                rep = None
            # 次要交付物：修正后的检查报告（生成失败不应阻断主交付物）
            if not annotate:
                try:
                    check_report = engine.run_check(odst, profile_path=profile)
                    engine.md_to_docx(check_report, chk)
                    self._debug(check_report)
                except Exception as e:
                    self._debug("[检查报告生成失败，已跳过] " + str(e))
                    chk = None
            outs[mode] = (odst, chk, rep)
        # 结果落到临时文件，交给主线程在「修正完成」后导出（先修正、后导出）
        self._fix_outs = outs

    # ----------------------------------------------------- 批量处理（导师版核心）
    def _batch_worker(self, idx, mode, out_dir):
        """第③步批量：将已选论文逐篇生成批注副本 / 一键修正，输出到同一文件夹，
        并生成一份汇总 Excel。画像只在首篇确认一次（避免逐个弹窗）。"""
        try:
            if not self._trial_ok():
                self.root.after(0, lambda: self._on_step_error(idx, mode, "试用次数已用完，请激活后使用"))
                return
            author = self.author_var.get().strip() or None
            modes = ["annotate", "fix"] if self._fix_mode == "both" else [self._fix_mode]
            mode_label = {"annotate": "生成批注副本",
                          "fix": "一键修正",
                          "both": "批注副本 + 已修正"}[self._fix_mode]
            # 批量只在首篇确认一次画像；后续沿用同一要求
            profile = self._confirm_profile_if_needed()
            # 注意：_confirm_profile_if_needed 在 worker 内用 Event 等待主线程弹窗，安全
            total = len(self.thesis_paths)
            rows = []
            group_cats = []
            for i, src in enumerate(self.thesis_paths, 1):
                self._set_status("正在处理第 %d/%d 篇：%s" % (i, total, os.path.basename(src)), RUN)
                # 同一时机刷新批量进度块（纯 UI 渲染，不影响下面的处理流程）
                self._batch_tick(i, total, os.path.basename(src))
                try:
                    docx_path, note = engine.normalize_input(src)
                except Exception as e:
                    rows.append([str(i), os.path.basename(src), mode_label, "（处理失败）", "错误：%s" % e])
                    continue
                if note:
                    self._debug(note)
                base = os.path.splitext(os.path.basename(src))[0]
                cr = None
                try:
                    if mode == "check":
                        # 检查模式批量：逐篇生成格式检查报告（.docx），不改动原文件
                        chk = os.path.join(out_dir, base + "_格式检查报告.docx")
                        cr = engine.run_check(docx_path, profile_path=profile)
                        engine.md_to_docx(cr, chk)
                        out_note = "检查报告：%s" % os.path.basename(chk)
                        rows.append([str(i), os.path.basename(src), "格式检查", out_note,
                                     "检查报告已生成"])
                    else:
                        for m in modes:
                            annotate = (m == "annotate")
                            suffix = "_批注副本" if annotate else "_已修正"
                            dst = os.path.join(out_dir, base + suffix + ".docx")
                            xlsx = os.path.join(out_dir, base + suffix + "_修改明细.xlsx")
                            chk = os.path.join(out_dir, base + suffix + "_检查报告.docx")
                            if annotate:
                                engine.run_annotate(docx_path, dst, profile_path=profile,
                                                    author=author, xlsx_path=xlsx)
                                out_note = "批注副本：%s" % os.path.basename(dst)
                            else:
                                rep = os.path.join(out_dir, base + suffix + "_修改报告.docx")
                                engine.run_fix_headings(docx_path, dst, profile_path=profile,
                                                        report_docx=rep, add_comments=True,
                                                        author=author, xlsx_path=xlsx)
                                try:
                                    cr = engine.run_check(dst, profile_path=profile)
                                    engine.md_to_docx(cr, chk)
                                except Exception as e:
                                    self._debug("[批量检查报告失败] " + str(e))
                                    chk = None
                                out_note = "修正稿：%s" % os.path.basename(dst)
                            rows.append([str(i), os.path.basename(src), mode_label, out_note,
                                         "明细见 %s" % os.path.basename(xlsx)])
                except Exception as e:
                    self._debug("[批量处理失败] %s → %s" % (src, e))
                    rows.append([str(i), os.path.basename(src), mode_label, "（处理失败）", "错误：%s" % e])
                # 全组共性问题：稿体检 → 归类 → 跨篇聚合（导师组会汇报用）。
                # 检查模式已算过 cr，直接复用；其余模式对原始稿体检。
                try:
                    _grp_md = cr if mode == "check" else engine.run_check(docx_path, profile_path=profile)
                    _grps = [classify_issue(m) for m in extract_issues_from_markdown(_grp_md)]
                    if _grps:
                        group_cats.append((os.path.basename(src), _grps))
                except Exception as e:
                    self._debug("[全组共性问题采集失败，已跳过] " + str(e))
            # 汇总 Excel（检查 / 修正 共用列表，仅标题与文件名按模式区分）
            _is_check = (mode == "check")
            summary = ["批量%s模式：%s" % ("检查" if _is_check else "处理", mode_label),
                       "共 %d 篇，署名：%s" % (total, author or "论文格式医生·导师版"),
                       ("每篇的检查报告已生成于本文件夹" if _is_check
                        else "每篇的逐条修改/批注明细见同名 _修改明细.xlsx")]
            columns = [("序号", 600), ("原文件名", 2600), ("交付方式", 1400),
                       ("输出文件", 2400), ("说明", 2600)]
            try:
                write_change_report_xlsx(
                    os.path.join(out_dir, "汇总_检查明细.xlsx" if _is_check else "汇总_修改明细.xlsx"),
                    ("批量检查汇总（%s）" if _is_check else "批量处理汇总（%s）")
                    % (author or "论文格式医生·导师版"), summary, columns, rows)
            except Exception as e:
                self._debug("[汇总 Excel 生成失败] " + str(e))
            # 全组共性问题总表（便于导师组会一句话汇报）
            try:
                _g_sum, _g_cols, _g_rows = aggregate(group_cats, total)
                write_change_report_xlsx(
                    os.path.join(out_dir, "汇总_全组共性问题.xlsx"),
                    "全组共性问题总表（%s）" % (author or "论文格式医生·导师版"), _g_sum, _g_cols, _g_rows)
            except Exception as e:
                self._debug("[全组共性问题总表生成失败] " + str(e))
            self._batch_out_dir = out_dir
            self.root.after(0, lambda: self._on_batch_done(idx, mode, out_dir, len(rows)))
        except Exception as e:
            self._debug("[批量错误] " + str(e))
            self.root.after(0, lambda: self._on_step_error(idx, mode, str(e)))
        finally:
            self.running = False
            engine.cleanup_conv_dirs()
            self.root.after(0, lambda: self._set_running(False))

    def _on_batch_done(self, idx, mode, out_dir, n):
        if mode == "check":
            # 检查模式批量完成：检查步骤（idx=1）已结束，推进到下一步（修正/交付方式）。
            self.step_index = idx + 1
            self._set_status("批量检查完成，已保存 %d 篇检查报告" % n, OKC)
            self._set_bar("done")
            self._refresh_wizard()
            self._show_batch_check_done(out_dir, n)
        else:
            # 批量修正完成：记录当前交付方式已导出（输出文件夹），切回该方式时按钮显示「已完成」
            self._fix_phase[self._fix_mode] = "saved"
            self._exported[self._fix_mode] = {out_dir}
            self._set_status("批量处理完成，已保存至文件夹", OKC)
            self._set_bar("done")
            self._refresh_wizard()
            self._show_batch_done(out_dir, n)

    def _show_batch_check_done(self, out_dir, n):
        msg = ("已批量检查 %d 篇论文，每篇的格式检查报告（_格式检查报告.docx）已保存在：\n%s\n\n"
               % (n, out_dir)
               + "另含一份「汇总_检查明细.xlsx」（逐篇结果）与「汇总_全组共性问题.xlsx」"
               "（共性问题汇总，便于导师组会一句话汇报）。\n\n"
               + "如需按学校要求进一步修正，请点「下一步」进入交付方式选择。")
        k = self._modal("批量检查完成", msg,
                        [("open", "打开输出文件夹"), ("ok", "完成")])
        if k == "open":
            self._open_folder(out_dir)

    def _show_batch_done(self, out_dir, n):
        msg = ("已批量处理 %d 篇论文，结果保存在：\n%s\n\n" % (n, out_dir)
               + "每篇均生成：论文（批注副本/修正稿）+ 同名 _修改明细.xlsx；\n"
               + "另含一份「汇总_修改明细.xlsx」便于整体核对。\n\n"
               + "全程在本机完成，原文件未改动。")
        k = self._modal("批量处理完成", msg,
                        [("open", "打开输出文件夹"), ("ok", "完成")])
        if k == "open":
            self._open_folder(out_dir)

    def _ask_trial_confirm(self, left):
        """修正前提示（worker 线程调用，主线程弹窗，Event 同步）。"""
        ev = threading.Event()
        box = {}

        def show():
            box["ok"] = messagebox.askyesno(
                "试用版提示",
                "当前为试用版（本机剩余 %d 次），修正后的论文将带水印，且为只读预览文档"
                "（编辑需密码）。\n\n"
                "正式版输出可编辑无水印文档，可一键交稿。\n\n是否继续？" % left,
                parent=self.root)
            ev.set()

        self.root.after(0, show)
        ev.wait(30)
        return bool(box.get("ok", False))

    def _show_trial_exhausted(self):
        """试用次数用完：弹购买引导（小程序码 + 去激活入口）。

        可能在 worker 线程也可能在**主线程**（root.after 排程）被调用：
          * worker 线程 → 用 Event 等主线程把弹窗弹完再返回；
          * 主线程     → **直接弹**，绝不能 ev.wait()：那会阻塞事件循环，
                         排队的 show 永远跑不到，界面要等超时才恢复（Codex 查出的冻结）。
        """
        text = ("本机免费试用（%d 次）已用完。\n\n"
                "正式版激活后：不限次数修正、输出无水印可编辑文档、一键交稿。\n\n"
                "微信扫下方小程序码即可购买激活码，付款后自动发码，立即可用。"
                % trial.TRIAL_LIMIT)
        if threading.current_thread() is threading.main_thread():
            self._show_purchase_qr("免费试用已用完", text)
            return
        ev = threading.Event()

        def show():
            try:
                self._show_purchase_qr("免费试用已用完", text)
            finally:
                ev.set()

        self.root.after(0, show)
        ev.wait(60)

    def _show_purchase_qr(self, title, text):
        """购买引导窗：展示小程序码（购买主渠道）+「我已购买 · 去激活」入口。

        必须在主线程调用（_show_trial_exhausted 已用 root.after(0) 调度）。
        小程序码缺失时降级为文字指引，绝不阻断主流程。
        """
        res = {"v": None}
        t = self._theme
        top = ModalShell(self.root, edition=self.edition, title=title, width=600, height=660)
        tk.Label(top.body, text=text, bg=t.modal_fill, fg=t.modal_sub,
                 font=t.sans(TYPE["body"]),
                 justify="left", wraplength=440).pack(padx=4, pady=(8, 10))
        shown = False
        try:
            if os.path.isfile(MINIAPP_QRCODE):
                img = tk.PhotoImage(file=MINIAPP_QRCODE)
                img = img.subsample(max(1, round(img.width() / 150)))
                lbl = tk.Label(top.body, image=img, bg=t.modal_fill)
                lbl.image = img          # 防 GC 回收导致图片不显示
                lbl.pack()
                tk.Label(top.body, text="微信扫一扫，或搜索【%s】小程序，付款后自动发码"
                         % MINIAPP_NAME, bg=t.modal_fill, fg=t.muted,
                         font=t.sans(TYPE["caption"])).pack(pady=(6, 0))
                shown = True
        except Exception:
            shown = False
        if not shown:
            tk.Label(top.body, text="（小程序码未随包提供，请前往官网 reedskill.com 购买）",
                     bg=t.modal_fill, fg=t.muted,
                     font=t.sans(TYPE["caption"])).pack(pady=(4, 0))
        fr = tk.Frame(top.body, bg=t.modal_fill)
        fr.pack(pady=(16, 4))
        widgets.RoundButton(fr, text="我已购买 · 去激活", edition=self.edition, style="primary",
                           height=SPACING["btn_h"], font=t.sans(TYPE["btn"], bold=True),
                           command=lambda: (res.update(v="activate"), top.destroy())).pack(side="left", padx=6)
        widgets.RoundButton(fr, text="稍后", edition=self.edition, style="secondary",
                           height=SPACING["btn_h"],
                           font=t.sans(TYPE["btn_small"], bold=True),
                           command=top.destroy).pack(side="left", padx=6)
        top.center_on(self.root)
        top.wait_window()
        if res["v"] == "activate":
            self._open_activation(force=True)

    def _show_upgrade_qr(self):
        """右上「升级正式版 →」入口：直接弹简单小程序码购买引导（老板拍板：不弹复杂激活窗 / 套餐卡片）。
        已购用户点「我已购买 · 去激活」仍走 show_activation 完整激活流程，激活能力不丢。
        """
        self._show_purchase_qr(
            "升级正式版",
            "正式版：不限次数修正、输出无水印可编辑文档、一键交稿。\n\n"
            "微信扫下方小程序码即可购买激活码，付款后自动发码，立即可用。")

    def _show_service_qr(self):
        """右上入口「客服微信」：沿用既有购买引导窗（客服微信 / 邮箱 + 小程序码 + 去激活）。"""
        self._show_purchase_qr(
            "联系客服",
            "客服微信：%s（ID：%s）\n邮箱：%s\n\n购买后自动发码；下方小程序码可直接扫码购买。"
            % (WECHAT_NAME, WECHAT_ID, ABOUT_MAIL))

    def _show_official_qr(self):
        """右上入口「官方公众号」：展示公众号二维码（复用既有 ModalShell，不另造弹窗样式）。"""
        t = self._theme
        top = ModalShell(self.root, edition=self.edition, title="官方公众号",
                         width=440, height=470)
        tk.Label(top.body,
                 text="微信扫码关注公众号【%s】（ID：%s）\n版本更新、使用技巧与激活教程都会发布在这里。"
                      % (WECHAT_NAME, WECHAT_ID),
                 bg=t.modal_fill, fg=t.modal_sub, font=t.sans(TYPE["body"]),
                 justify="left", wraplength=360).pack(padx=4, pady=(8, 10))
        shown = False
        try:
            if os.path.isfile(QRCODE):
                img = tk.PhotoImage(file=QRCODE)
                img = img.subsample(max(1, round(img.width() / 150)))
                lbl = tk.Label(top.body, image=img, bg=t.modal_fill)
                lbl.image = img          # 防 GC 回收导致图片不显示
                lbl.pack()
                shown = True
        except Exception:
            shown = False
        if not shown:
            tk.Label(top.body, text="（二维码未随包提供，请前往官网 reedskill.com）",
                     bg=t.modal_fill, fg=t.muted,
                     font=t.sans(TYPE["caption"])).pack(pady=(6, 0))
        widgets.RoundButton(top.body, text="关闭", edition=self.edition, style="secondary",
                            height=SPACING["btn_h"],
                            font=t.sans(TYPE["btn_small"], bold=True),
                            command=top.destroy).pack(pady=(18, 0))
        top.center_on(self.root)
        top.wait_window()

    def _on_trial_blocked(self):
        """试用被拦截（用完/取消）：停留在当前步，给友好提示。"""
        self._errored = False
        # 用户刚在「购买引导」里点「去激活」并激活成功时，worker 仍会返回 False 走到这里；
        # 此时不能再报「试用已用完」，否则与刚写入的「已激活正式版」自相矛盾（Codex N1）。
        try:
            if trial.is_licensed():
                self._update_trial_badge()
                self._set_status("已激活正式版，感谢支持", OKC)
                self._set_bar("idle")
                self._refresh_wizard()
                return
        except Exception:
            pass
        self._set_status("试用次数已用完，请激活后使用", ERRC)
        self._set_bar("idle")
        self._refresh_wizard()

    # ------------------------------------------------------------ 确认页
    def _show_check_done(self, out_md):
        k = self._modal("格式检查完成",
                        "检查报告已保存至：\n%s\n\n如需按学校要求修正论文，请继续第③步。" % out_md,
                        [("open", "打开所在文件夹"), ("ok", "完成")])
        if k == "open":
            self._open_folder(out_md)

    def _show_fix_done(self, dst, chk, rep):
        lines = ["已为您保存以下文件：\n"]
        lines.append("① 修正后论文：%s" % dst)
        lines.append("② 修改报告：%s" % rep)
        lines.append("③ 检查报告：%s" % (chk if chk else "（本次未生成，可点“再处理一篇”重试）"))
        lines.append("\n全程在本机完成，原文件未改动。")
        k = self._modal("修正完成", "\n".join(lines),
                        [("open", "打开所在文件夹"), ("ok", "完成")])
        if k == "open":
            self._open_folder(dst)

    def _export_fix(self):
        """主线程：修正完成后让客户选择保存位置并导出三个文件（先修正、后导出）。

        同一会话内已导出过则先弹“已保存过，是否再次保存（覆盖）”。
        """
        outs = getattr(self, "_fix_outs", None) or {}
        if not outs:
            return
        modes = sorted(outs.keys())
        # v1.0.15：按当前交付方式判断是否已导出过（此前为全局 _fix_saved，切模式后会误判）
        prev = (self._exported or {}).get(self._fix_mode)
        if prev:
            if not messagebox.askyesno(
                    "已导出过",
                    "该交付方式之前已导出过：\n%s\n\n确定要再次生成并保存（覆盖）吗？"
                    % "\n".join(sorted(os.path.basename(p) for p in prev))):
                return
        base_src = _base_no_ext(self.thesis_path.get().strip() or "论文")
        saved = set()
        temp_dirs = set()

        def _copy_set(tmp_dst, chk, rep, dst):
            """把一套产出（论文/报告/检查/明细）复制到 dst，并记录已保存与临时目录。"""
            base = _base_no_ext(dst)
            shutil.copy(tmp_dst, dst)
            saved.add(dst)
            if rep:
                shutil.copy(rep, base + "_修改报告.docx")
                saved.add(base + "_修改报告.docx")
            if chk:
                shutil.copy(chk, base + "_检查报告.docx")
                saved.add(base + "_检查报告.docx")
            _src_xlsx = os.path.splitext(tmp_dst)[0] + "_修改明细.xlsx"
            if os.path.isfile(_src_xlsx):
                shutil.copy(_src_xlsx, base + "_修改明细.xlsx")
                saved.add(base + "_修改明细.xlsx")
            temp_dirs.add(os.path.dirname(tmp_dst))

        if len(modes) == 1:
            m = modes[0]
            tmp_dst, chk, rep = outs[m]
            _suffix = "_批注副本" if m == "annotate" else "_已修正"
            dst = filedialog.asksaveasfilename(
                title="选择保存位置",
                initialfile=os.path.basename(base_src) + _suffix + ".docx",
                initialdir=os.path.dirname(base_src) or None,
                defaultextension=".docx",
                filetypes=[("Word 文档", "*.docx")])
            if not dst:
                messagebox.showinfo("未导出",
                                    "未选择保存位置，本次结果未导出。\n"
                                    "如需导出，请点“上一步”回到第③步重新执行。")
                return
            try:
                _copy_set(tmp_dst, chk, rep, dst)
            except Exception as e:
                messagebox.showerror("导出失败", str(e) + HELP_HINT)
                return
            base = _base_no_ext(dst)
            self._finish_export(saved, temp_dirs, dst,
                                (base + "_检查报告.docx") if chk else None,
                                (base + "_修改报告.docx") if rep else None)
            return

        # ①+② 都要：弹文件夹，两套一起导出
        out_dir = filedialog.askdirectory(title="选择保存文件夹（将同时导出批注副本与已修正版）")
        if not out_dir:
            messagebox.showinfo("未导出",
                                "未选择保存文件夹，本次结果未导出。\n"
                                "如需导出，请点“上一步”回到第③步重新执行。")
            return
        try:
            for m in modes:
                tmp_dst, chk, rep = outs[m]
                suffix = "_批注副本" if m == "annotate" else "_已修正"
                dst = os.path.join(out_dir, os.path.basename(base_src) + suffix + ".docx")
                _copy_set(tmp_dst, chk, rep, dst)
        except Exception as e:
            messagebox.showerror("导出失败", str(e) + HELP_HINT)
            return
        self._finish_export(saved, temp_dirs,
                            os.path.join(out_dir, os.path.basename(base_src) + "_已修正.docx"),
                            None, None)

    def _finish_export(self, saved, temp_dirs, dst, chk_path, rep_path):
        """收尾：记录已保存、清理临时目录、进入“完成”态并弹完成提示。

        v1.0.15：saved 按当前交付方式记录（_fix_phase[mode] + _exported[mode]），
        三个交付方式可各自完成导出互不覆盖；切回已导出的方式时按钮显示「已完成」。
        """
        self._fix_saved = saved
        try:
            for d in temp_dirs:
                shutil.rmtree(d, ignore_errors=True)
        except Exception:
            pass
        self._fix_phase[self._fix_mode] = "saved"
        self._exported[self._fix_mode] = set(saved)
        self._set_status("已保存，可再处理一篇", OKC)
        self._refresh_wizard()
        self._show_fix_done(dst, chk_path, rep_path)

    def _modal(self, title, text, buttons):
        result = {"v": None}
        t = self._theme
        top = ModalShell(self.root, edition=self.edition, title=title, width=620, height=320)
        tk.Label(top.body, text=text, bg=t.modal_fill, fg=t.modal_sub,
                 font=t.sans(TYPE["body"]), justify="left", wraplength=520).pack(
            anchor="w", pady=(6, 24))
        fr = tk.Frame(top.body, bg=t.modal_fill)
        fr.pack(fill="x", pady=(0, 6))
        for key, label in buttons:
            st = "primary" if key == "ok" else "secondary"
            b = widgets.RoundButton(fr, text=label, edition=self.edition, style=st,
                                   height=SPACING["btn_h"],
                                   font=t.sans(TYPE["btn"] if key == "ok" else TYPE["btn_small"],
                                               bold=True),
                                   command=lambda k=key: (result.update(v=k), top.destroy()))
            b.pack(side="left", padx=(0, 10))
        top.center_on(self.root)
        top.wait_window()
        return result["v"]

    def _open_folder(self, path):
        folder = os.path.dirname(os.path.abspath(path))
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception:
            pass

    # ------------------------------------------------- 桌面快捷方式（v1.0.31 绿色 zip 自建）
    def _desktop_dir(self):
        """真实桌面路径（兼容 OneDrive 重定向）；拿不到时退回 ~/Desktop。"""
        try:
            import win32com.client as _wc
            ws = _wc.Dispatch("WScript.Shell")
            d = ws.SpecialFolders("Desktop")
            if d:
                return d
        except Exception:
            pass
        return os.path.join(os.path.expanduser("~"), "Desktop")

    def _desktop_shortcut_path(self):
        return os.path.join(self._desktop_dir(), APP_SHORTCUT_NAME + ".lnk")

    def _desktop_shortcut_exists(self):
        try:
            return os.path.isfile(self._desktop_shortcut_path())
        except Exception:
            return False

    def _create_desktop_shortcut(self, silent=False):
        """创建桌面快捷方式（指向当前主程序）。成功 True；silent=True 时静默不弹提示。"""
        try:
            if not (sys.platform.startswith("win") and _is_frozen_exe()):
                if not silent:
                    messagebox.showinfo("桌面快捷方式",
                                        "仅 Windows 正式版支持自动创建；\n"
                                        "可手动：右键主程序 → 发送到 → 桌面快捷方式。",
                                        parent=self.root)
                return False
            import win32com.client as _wc
            ws = _wc.Dispatch("WScript.Shell")
            exe = os.path.abspath(sys.executable)
            lnk = self._desktop_shortcut_path()
            sc = ws.CreateShortcut(lnk)
            sc.TargetPath = exe
            sc.WorkingDirectory = os.path.dirname(exe)
            sc.IconLocation = exe + ",0"
            sc.Save()
            if not silent:
                messagebox.showinfo("桌面快捷方式",
                                    "已创建「%s」桌面快捷方式，双击即可打开。" % APP_SHORTCUT_NAME,
                                    parent=self.root)
            return True
        except Exception as e:
            if not silent:
                messagebox.showerror("创建失败",
                                     "自动创建失败：%s\n\n"
                                     "请手动：右键主程序 → 发送到 → 桌面快捷方式。" % e,
                                     parent=self.root)
            return False

    def _maybe_auto_shortcut(self):
        """首次启动自动在桌面创建一次快捷方式（仅 Windows 正式版、桌面尚无该图标时）。
        静默执行，不弹窗；已创建过则不再重复（标记文件存 %APPDATA%\\SHORTCUT_FLAG_TAG）。
        顶栏“桌面图标”链接可随时手动补建 / 重建。"""
        try:
            if not (sys.platform.startswith("win") and _is_frozen_exe()):
                return
            if self._desktop_shortcut_exists():
                return
            flag = os.path.join(os.environ.get("APPDATA", ""), SHORTCUT_FLAG_TAG,
                                "shortcut_auto.txt")
            if os.path.isfile(flag):
                return
            os.makedirs(os.path.dirname(flag), exist_ok=True)
            with open(flag, "w", encoding="utf-8") as f:
                f.write("1")
            self._create_desktop_shortcut(silent=True)
        except Exception:
            pass

    def _update_profile_box(self):
        p = self.profile_path.get().strip()
        if p and os.path.isfile(p):
            self.profile_info_var.set("已载入格式画像：" + os.path.basename(p))
            if not self.profile_box.winfo_ismapped():
                self.profile_box.pack(fill="x", pady=(10, 0), after=self._template_box)
        else:
            if self.profile_box.winfo_ismapped():
                self.profile_box.pack_forget()

    def _reset_wizard(self):
        """“再处理一篇”：回到第 1 步并清空选择。"""
        self.step_index = 0
        self._errored = False
        self._profile_confirmed = False
        self._profile_abandoned = False
        self._check_report_saved = None
        self._fix_saved = set()
        self._fix_outs = {}
        self._fix_phase = {}
        self._exported = {}
        self.thesis_paths = []
        self._fix_mode = "annotate"
        self._fix_mode_var.set("annotate")
        self.thesis_path.set("")
        self.template_path.set("")
        self.profile_path.set("")
        self._template_name.config(text="用于按学校要求检查／修正，更贴合要求", fg=MUTED)
        self._thesis_dot.config(text="○", fg=MUTED)
        self._update_thesis_status()
        self._tpl_dot.config(text="○", fg=MUTED)
        self._tpl_lbl.config(text="模板未选（可选）", fg=MUTED)
        self._update_profile_box()
        self._set_status("请按步骤操作", MUTED)
        self._set_bar("idle")
        self._refresh_wizard()
        self._paint_mode_buttons()
        # 已记住的模板：再处理一篇时自动重新载入，免得重复选
        if self._tpl_locked:
            self._restore_template_lock()

    # ---------------------------------------------- 内部日志（不展示客户）
    def _debug(self, text):
        self._msgs.append(text)
        if len(self._msgs) > 200:
            self._msgs = self._msgs[-200:]

    def _set_status(self, s, color):
        self.root.after(0, lambda: (self.status_var.set(s),
                                    self.status_lbl.config(fg=color),
                                    self.status_dot.config(fg=color)))

    def _update_trial_badge(self):
        """刷新顶部入口行的试用 / 正式标识（只换渲染对象，触发时机与判定一字未动）。"""
        try:
            self._licensed = trial.is_licensed()
        except Exception:
            self._licensed = False
        if self._licensed:
            # v1.0.23：区分卡种展示——周卡/月卡/次卡不再是"永久"，要让用户看得见期限与余量
            try:
                s = license.license_summary()
            except Exception:
                s = None
            self._trial_badge.set_kind("formal")
            # badge 形如「正式版 · 周卡（剩 5 天）」，本身就是完整短文案，不再另加前缀
            self._trial_badge.set_text(s["badge"] if s else "正式版")
            self._upgrade_btn.config(text="已激活 ✓", state="disabled")
        else:
            self._trial_badge.set_kind("trial")
            left = trial.trials_left()
            if left > 0:
                self._trial_badge.set_text("试用版 · 剩余 %d 次" % left)
            else:
                self._trial_badge.set_text("试用版 · 已用完")
            self._upgrade_btn.config(text="升级正式版 →", state="normal")

    def _open_activation(self, force: bool = False):
        """主界面右上角激活入口：打开激活窗，成功后刷新状态。

        ``force=True``：来自「试用用完 → 购买引导」弹窗的「我已购买 · 去激活」。
        此时 worker 线程仍在跑（`self.running` 要等 _do_fix 返回才清），走 running 守卫
        会直接弹「正在处理」→ 激活窗永远打不开（Codex 2026-09-12 审查发现的 P1）。
        """
        if self.running and not force:
            messagebox.showinfo("正在处理", "上一步还在处理中，请稍候…", parent=self.root)
            return
        r = show_activation(self.root, show_trial=False)
        if r == "ok":
            self._update_trial_badge()
            try:
                s = license.license_summary()
            except Exception:
                s = None
            self._set_status(("已激活%s · %s" % (s["kind"], s["desc"])) if s
                             else "已激活正式版，感谢支持", OKC)
        elif r == "trial":
            self._update_trial_badge()
            self._set_status("已进入试用模式", MUTED)

    def _handle_revoked(self):
        """授权被后台撤销（退款锁死）：锁定软件，禁止继续修正。"""
        self._licensed = False
        self._update_trial_badge()
        self._set_status("授权已失效（可能已退款），软件已锁定", ERRC)
        self._modal(
            "授权已失效",
            "您的授权已被后台撤销（可能因退款）。\n\n"
            "软件已锁定，无法继续修正论文。\n如需继续使用，请联系客服：hi@reedskill.com。",
            [("ok", "知道了")])


# ---------------------------------------------------------------------------
_APP_REF = None  # App 实例引用，在 App.__init__ 中赋值


def _on_license_revoked():
    """心跳线程发现 revoked：通过 after(0) 切回主线程处理，避免后台线程碰 Tk。"""
    if _APP_REF is not None:
        try:
            _APP_REF.root.after(0, _APP_REF._handle_revoked)
        except Exception:
            pass


def show_activation(root, show_trial=True):
    """激活窗口：在线激活（主） + 离线备用码（兜底）。

    show_trial: 是否显示"先试用"入口。仅启动时首次弹窗为 True；
    从主界面激活入口打开时客户已在试用模式，无需再显示（v1.3.66）。

    返回：
      "ok"     激活成功；
      "trial"  客户选择"先试用"（暂不激活，进入试用模式）；
      "quit"   直接关闭窗口（退出程序）。
    """
    result = {"v": "quit"}
    t = get_theme("advisor")
    _w, _h = 620, 740
    top = ModalShell(root, edition="advisor", title="导师版 · 升级正式版",
                    width=_w, height=_h)

    tk.Label(top.body, text="论文格式医生", bg=t.modal_fill, fg=t.modal_title,
             font=t.serif(TYPE["page_title"], bold=True)).pack(pady=(14, 4))
    # v1.0.23：卡片种类已扩展到 4 种（永久 / 周卡 / 月卡 / 次卡），文案不再统一说"永久、完全离线"
    tk.Label(top.body,
             text="请输入您购买的激活码以激活。\n"
                  "永久卡：激活后完全离线、永久可用；周卡 / 月卡：期限内可用；\n"
                  "次卡：每次修正需联网扣一次次数，离线时不可使用。",
             bg=t.modal_fill, fg=t.modal_sub, font=t.sans(TYPE["caption"]),
             wraplength=440, justify="center").pack(pady=(0, 10))

    card_var = tk.StringVar()
    code_entry = tk.Entry(top.body, textvariable=card_var,
                          font=t.sans(TYPE["body"]), width=40,
                          relief="solid", bd=1, highlightthickness=1,
                          highlightbackground=t.select_border, bg=t.surface,
                          justify="center")
    code_entry.pack(pady=(4, 6))
    # 占位文案照 COPY_TABLE.md 第 5 节：`输入激活码`（tk 没有原生 placeholder，自管两态）
    card_var.set(widgets.COPY["input_code_ph"])
    code_entry.configure(fg=t.muted)

    def _code_focus_in(_e=None):
        if card_var.get() == widgets.COPY["input_code_ph"]:
            card_var.set("")
            code_entry.configure(fg=t.modal_title)

    def _code_focus_out(_e=None):
        if not card_var.get().strip():
            card_var.set(widgets.COPY["input_code_ph"])
            code_entry.configure(fg=t.muted)

    code_entry.bind("<FocusIn>", _code_focus_in)
    code_entry.bind("<FocusOut>", _code_focus_out)

    msg_var = tk.StringVar()
    tk.Label(top.body, textvariable=msg_var, bg=t.modal_fill, fg=t.error_text,
             font=t.sans(TYPE["caption"])).pack(pady=(0, 6))

    def do_activate():
        card = card_var.get().strip()
        if not card:
            msg_var.set("请输入激活码")
            return
        btn_activate.config(state="disabled")
        msg_var.set("正在验证您的授权，请稍候…（首次激活需联网校验，通常需要 20~40 秒）")
        q = queue.Queue()
        start_t = time.time()

        def work():
            # 工作线程只做两件事：跑验证 + 把结果放进队列；绝不直接碰 Tk 界面
            try:
                mc = license.get_machine_code()
                license._log("gui: 开始联网验证 card=%s mc=%s" % (card, mc))
                ok, _days, note = license.verify_via_kami(card, mc)
                license._log("gui: 联网验证返回 ok=%s note=%s" % (ok, note))
                q.put(("done", ok, note, mc))
            except Exception:
                import traceback
                tb = traceback.format_exc()
                license._log("gui: 激活线程异常\n%s" % tb)
                q.put(("done", False, "激活过程出错，请重试或联系客服", ""))

        threading.Thread(target=work, daemon=True).start()

        def poll():
            # 主线程轮询队列（线程安全），拿到结果才更新界面
            try:
                _kind, ok, note, mc = q.get_nowait()
            except queue.Empty:
                if time.time() - start_t > 120:   # UI 看门狗：物理上不可能无限转圈
                    btn_activate.config(state="normal")
                    msg_var.set("验证超时：请检查网络后重试；或联系客服获取离线激活码")
                    license._log("gui: UI 看门狗触发（120 秒未等到结果）")
                    return
                top.after(300, poll)
                return
            btn_activate.config(state="normal")
            if ok:
                license.save_local_license(card, mc)
                license._log("gui: 授权已写入本机")
                # 启动后台心跳：联网时若授权被撤销（退款），自动锁死
                license.start_heartbeat(
                    product=license.DEFAULT_PRODUCT, on_revoked=_on_license_revoked)
                result["v"] = "ok"
                top.destroy()
            else:
                msg_var.set(note)

        top.after(300, poll)

    def _clear_code_placeholder():
        if card_var.get() == widgets.COPY["input_code_ph"]:
            card_var.set("")

    btn_activate = widgets.RoundButton(top.body, text="激活", edition="advisor", style="primary",
                                      height=SPACING["btn_h"],
                                      font=t.sans(TYPE["btn"], bold=True),
                                      command=lambda: (_clear_code_placeholder(), do_activate()))
    btn_activate.pack(pady=(2, 4))

    # v1.3.58：试用入口放在显眼位置（激活按钮正下方，便于未购买客户先体验）
    # v1.3.66：仅启动时首次弹窗显示；从主界面激活入口打开时客户已在试用模式，无需再显示
    if show_trial:
        widgets.RoundButton(top.body,
                           text="还没有激活码？先试用（免费 %d 次）" % trial.TRIAL_LIMIT,
                           edition="advisor", style="secondary", height=SPACING["btn_h"],
                           font=t.sans(TYPE["btn_small"], bold=True),
                           command=lambda: (result.update(v="trial"), top.destroy())).pack(pady=(4, 2))
        tk.Label(top.body, text="试用版可完整体验一键修正，输出带水印且为只读预览；正式版可编辑无水印。",
                 bg=t.modal_fill, fg=t.muted, font=t.sans(TYPE["caption"]),
                 wraplength=540).pack(pady=(0, 6))

    tk.Label(top.body, text="— 以下为特殊情形使用 —", bg=t.modal_fill, fg=t.muted,
             font=t.sans(TYPE["caption"])).pack(pady=(8, 4))
    mc = license.get_machine_code()
    # 机器码 + 复制按钮并排一行（省高度）
    mc_row = tk.Frame(top.body, bg=t.modal_fill)
    mc_row.pack(pady=(0, 6))
    tk.Label(mc_row, text="本机机器码：" + mc, bg=t.modal_fill, fg=t.muted,
             font=t.mono(TYPE["mono"])).pack(side="left")
    widgets.RoundButton(mc_row, text="复制", edition="advisor", style="secondary",
                       height=SPACING["btn_h_sm"],
                       font=t.sans(TYPE["btn_small"], bold=True),
                       command=lambda: top.clipboard_append(mc)).pack(side="left", padx=(8, 0))

    off_var = tk.StringVar()
    tk.Label(top.body, text="离线激活码（网络不通时，联系客服获取）：", bg=t.modal_fill, fg=t.muted,
             font=t.sans(TYPE["caption"])).pack(pady=(2, 2))
    tk.Entry(top.body, textvariable=off_var, width=46, font=t.mono(TYPE["mono"]),
             relief="solid", bd=1, highlightthickness=1,
             highlightbackground=t.select_border, bg=t.surface).pack(pady=(2, 4))

    def do_offline():
        code = off_var.get().strip()
        if not code:
            msg_var.set("请输入离线激活码")
            return
        mc2 = license.get_machine_code()
        if license.verify_offline_code(code, mc2):
            license.save_offline_license(code, mc2)
            result["v"] = "ok"
            top.destroy()
        else:
            msg_var.set("离线激活码无效，请核对后重试")

    widgets.RoundButton(top.body, text="使用离线激活码激活", edition="advisor", style="secondary",
                       height=SPACING["btn_h"],
                       font=t.sans(TYPE["btn_small"], bold=True),
                       command=do_offline).pack(pady=(2, 4))

    widgets.RoundButton(top.body, text="退出", edition="advisor", style="secondary",
                       height=SPACING["btn_h"],
                       font=t.sans(TYPE["btn_small"], bold=True),
                       command=lambda: top.destroy()).pack(pady=(0, 4))

    hl = tk.Frame(top.body, bg=t.modal_fill)
    hl.pack(pady=(4, 2))
    tk.Button(hl, text="关于", bg=t.modal_fill, fg=t.primary,
              font=t.sans(TYPE["caption"]),
              relief="flat", cursor="hand2",
              command=lambda: show_about(top)).pack(side="left", padx=14)
    tk.Button(hl, text="使用帮助", bg=t.modal_fill, fg=t.primary,
              font=t.sans(TYPE["caption"]),
              relief="flat", cursor="hand2",
              command=lambda: show_help(top)).pack(side="left", padx=14)
    _site_label(hl, bg=t.modal_fill).pack(side="left", padx=14)

    tk.Label(top.body, text="未签名程序提示：Windows 可能弹出 SmartScreen 拦截，点击「详细信息」→「仍要运行」即可打开（官网激活教程有图文演示）。",
             bg=t.modal_fill, fg=t.muted, font=t.sans(TYPE["caption"]),
             wraplength=560).pack(pady=(2, 6))
    tk.Label(top.body, text="客服微信 · 官方公众号 · hi@reedskill.com",
             bg=t.modal_fill, fg=t.muted, font=t.sans(TYPE["footer"]),
             wraplength=560).pack(pady=(8, 10))

    # v1.3.67：按内容实际所需高度自动伸缩窗口（Tk 自动测量，保证底部版权完整显示不截断）
    try:
        top.update_idletasks()
        _req_h = top.inner.winfo_reqheight()
        _req_w = top.inner.winfo_reqwidth()
        _sh = top.winfo_screenheight()
        _sw = top.winfo_screenwidth()
        top._width = min(max(_req_w + 40, 600), _sw - 60)
        top._height = min(max(_req_h + 40, 720), _sh - 80)
    except Exception:
        pass
    top.center_on(root)
    top.wait_window()
    return result["v"]


def _site_label(parent, text=OFFICIAL_SITE_TEXT, bg=PAPER):
    """可点击的官网链接标签：点一下用默认浏览器打开官网。"""
    lbl = tk.Label(parent, text=text, bg=bg, fg="#185FA5", font=F_SMALL, cursor="hand2")
    lbl.bind("<Button-1>", lambda e: webbrowser.open(OFFICIAL_SITE))
    lbl.bind("<Enter>", lambda e: lbl.config(fg="#0c447c"))
    lbl.bind("<Leave>", lambda e: lbl.config(fg="#185FA5"))
    return lbl


# ---------------------------------------------------------------------------
# 关于 / 帮助 窗口（v1.3.68：品牌署名 + 客服入口 + 主打论文安全·离线）
# ---------------------------------------------------------------------------
def _scroll_frame(parent, bg=PAPER):
    """可滚动容器：Canvas + 常驻滚动条，滚轮/拖拽/箭头均可用，内部宽度跟随画布。"""
    cv = tk.Canvas(parent, bg=bg, highlightthickness=0)
    sb = ttk.Scrollbar(parent, orient="vertical", command=cv.yview)
    inner = tk.Frame(cv, bg=bg)
    win = cv.create_window((0, 0), window=inner, anchor="nw")

    def _sync(event=None):
        # 先让内部布局稳定（wraplength 重排等），再按实际内容刷新滚动区域；
        # 否则 bbox 高度滞后，滑块会显示成满格"一条"、拖不动（只能点箭头）。
        try:
            cv.update_idletasks()
        except Exception:
            pass
        cv.configure(scrollregion=cv.bbox("all"))

    def _on_cv(event):
        cv.itemconfig(win, width=event.width)
        _sync()

    def _wheel(event):
        if getattr(event, "num", None) == 4:
            cv.yview_scroll(-3, "units")
        elif getattr(event, "num", None) == 5:
            cv.yview_scroll(3, "units")
        else:
            d = getattr(event, "delta", 0) or 0
            step = -3 if d > 0 else 3
            if abs(d) >= 120:
                step = -int(d / 120) * 3
            cv.yview_scroll(step, "units")

    inner.bind("<Configure>", _sync)
    cv.bind("<Configure>", _on_cv)
    cv.configure(yscrollcommand=sb.set)
    cv.grid(row=0, column=0, sticky="nsew")
    sb.grid(row=0, column=1, sticky="ns")
    parent.rowconfigure(0, weight=1)
    parent.columnconfigure(0, weight=1)
    # 兜底：窗口显示、内容布局稳定后多次刷新滚动范围（滑块大小/上下界限由此决定）
    parent.after(150, _sync)
    parent.after(500, _sync)
    for w in (cv, inner):
        w.bind("<MouseWheel>", _wheel)
        w.bind("<Button-4>", _wheel)
        w.bind("<Button-5>", _wheel)
    return inner


def _auto_wrap(lbl, container, padx, extra=0):
    """让 Label 的换行宽度跟随容器宽度，窗口缩放时文字自动重新排版。"""
    def _fit(event):
        w = event.width - 2 * padx - extra
        if w > 80:
            lbl.configure(wraplength=w)
    container.bind("<Configure>", _fit)
    return lbl


def show_about(parent):
    """关于页：名片式居中排版，细线分隔，文字随窗口自适应。"""
    top = tk.Toplevel(parent)
    top.title("关于 · 论文格式医生·导师版")
    top.configure(bg=PAPER)
    top.geometry("565x680")
    top.resizable(True, True)
    top.minsize(520, 560)

    inner = _scroll_frame(top)
    padx = 44

    def wlbl(text, fg, font, pady, nowrap=False, **kw):
        """居中文本标签；nowrap=True 固定单行（短信息），否则换行宽度跟随窗口。"""
        opts = dict(bg=PAPER, fg=fg, font=font, justify="center")
        if not nowrap:
            opts["wraplength"] = 420
        lbl = tk.Label(inner, text=text, **opts)
        if not nowrap:
            _auto_wrap(lbl, inner, padx)
        lbl.pack(padx=padx, pady=pady)
        return lbl

    def rule(pady=(0, 14)):
        tk.Frame(inner, bg=LINE, height=1).pack(fill="x", padx=padx, pady=pady)

    # —— 头部：题名 + 版本 ——
    tk.Label(inner, text="论文格式医生", bg=PAPER, fg=INK,
             font=F_TITLE).pack(padx=padx, pady=(24, 2))
    tk.Label(inner, text="版本 v%s" % APP_VERSION, bg=PAPER, fg=MUTED,
             font=F_SUB).pack(padx=padx, pady=(0, 12))
    rule()

    # —— 定位 ——
    wlbl("以学校模板为准绳，为论文格式把脉。", INK, F_HDR, (0, 8), nowrap=True)
    wlbl("标题层级、字体字号、页边距、行距，参考文献与三线表，逐一对照，改至合乎规范。",
         BODY, F_BODY, (0, 14))
    rule()

    # —— 品牌 + 关注 ——
    wlbl("芦苇不熬夜 出品", INK, F_HDR, (2, 8), nowrap=True)
    try:
        if os.path.isfile(QRCODE):
            qr = tk.PhotoImage(file=QRCODE)
            qr = qr.subsample(max(1, round(qr.width() / 150)))  # 缩到约 150px
            ql = tk.Label(inner, image=qr, bg=PAPER)
            ql.image = qr
            ql.pack(pady=(0, 8))
    except Exception:
        pass
    wlbl("微信公众号：【%s】（ID：%s）" % (WECHAT_NAME, WECHAT_ID), BODY, F_BODY, (0, 2), nowrap=True)
    wlbl("联系邮箱：%s" % ABOUT_MAIL, BODY, F_BODY, (0, 6), nowrap=True)
    wlbl("使用中若有疑问，欢迎关注公众号留言，或致信 %s，我们看到即复。" % ABOUT_MAIL,
         MUTED, F_SMALL, (0, 14))
    _site_label(inner, text="官网 reedskill.com · 产品介绍与激活教程").pack(padx=padx, pady=(0, 6))
    rule()

    # —— 关于本软件 ——
    wlbl("关于本软件", ACCENT, F_HDR, (0, 8))
    wlbl("论文安全 · 所有修正均在本机完成，论文不联网、不上传、不被收集。",
         BODY, F_SMALL, (0, 4))
    wlbl("模板驱动 · 格式要求取自学校模板，批注说明为先，样式定义次之。",
         BODY, F_SMALL, (0, 12))
    wlbl("未签名程序 · 首次打开时 Windows 可能弹出 SmartScreen 拦截，点击「详细信息」→「仍要运行」即可（官网激活教程有图文演示）。",
         MUTED, F_SMALL, (0, 12))

    tk.Label(inner, text="© 2026 芦苇不熬夜", bg=PAPER, fg=MUTED,
             font=F_FOOT).pack(padx=padx, pady=(0, 18))


def show_help(parent):
    """使用帮助：文档式左对齐排版，分区细线，文字随窗口自适应。"""
    top = tk.Toplevel(parent)
    top.title("使用帮助 · 论文格式医生·导师版")
    top.configure(bg=PAPER)
    top.geometry("805x680")
    top.resizable(True, True)
    top.minsize(680, 560)

    inner = _scroll_frame(top)
    padx = 40
    wrap = 660

    def albl(text, fg, font, nowrap=False, **kw):
        """左对齐文本标签；nowrap=True 固定单行，否则换行宽度跟随窗口。"""
        opts = dict(bg=PAPER, fg=fg, font=font, justify="left", anchor="w")
        if not nowrap:
            opts["wraplength"] = wrap
        lbl = tk.Label(inner, text=text, **opts)
        if not nowrap:
            _auto_wrap(lbl, inner, padx)
        return lbl

    # —— 头部（居中）——
    tk.Label(inner, text="使用帮助", bg=PAPER, fg=INK,
             font=F_TITLE).pack(padx=padx, pady=(22, 4))
    tk.Label(inner, text="将学校模板告知软件，导入论文后依向导循序而行，格式自可妥帖。",
             bg=PAPER, fg=BODY, font=F_BODY, justify="center").pack(padx=padx, pady=(0, 12))
    tk.Frame(inner, bg=LINE, height=1).pack(fill="x", padx=padx, pady=(0, 8))

    def section(title):
        tk.Label(inner, text=title, bg=PAPER, fg=ACCENT, font=F_HDR,
                 anchor="w").pack(fill="x", padx=padx, pady=(14, 4))
        tk.Frame(inner, bg=LINE, height=1).pack(fill="x", padx=padx, pady=(0, 4))

    def item(title, body):
        tl = albl(title, INK, _ADV.serif(TYPE["card_title"], bold=True))
        tl.pack(fill="x", padx=padx, pady=(8, 1))
        bl = albl(body, BODY, F_BODY)
        bl.pack(fill="x", padx=padx, pady=(0, 4))

    # —— 一、怎么用 ——
    section("一、怎么用")
    steps = [
        ("1. 导入论文",
         "在主界面「文件选择」中导入待处理的论文（.docx / .doc / .wps）。"
         "有两个按钮：\n"
         "·「选择多个文件」：逐个或按住 Ctrl 多选文件；\n"
         "·「选择文件夹」：若学生论文都放在同一个文件夹里，点它即可自动收齐该文件夹内的所有 Word 文档，无需一篇篇选。\n"
         "下方会实时显示「已选择 N 篇论文」。导入后软件进入向导模式，引领你完成后续各步。"),
        ("2. 导入学校模板",
         "仍在「文件选择」处，选择学校下发的格式要求文件（Word 模板或格式规范文档均可）。"
         "模板若带批注更佳：批注常写明具体要求，如「一级标题用黑体小二、居中」；"
         "软件以批注说明为先，故带批注者修正最贴近学校要求。"),
        ("3. 生成模板画像",
         "软件读取模板中的格式要求，生成「模板画像」——字体、字号、行距、页边距、标题层级等要素。"),
        ("4. 逐步确认",
         "软件分步呈现各项设定（标题层级、字体字号、页边距、缩进、行距、参考文献、三线表），"
         "逐一过目、确认即可。"),
        ("5. 生成报告",
         "确认完毕后，软件依模板修正论文，并生成报告，注明改动之处与未能确定之处。"),
    ]
    for title, body in steps:
        item(title, body)

    # —— 二、常见问题 ——
    section("二、常见问题")
    faqs = [
        ("Q1 · 什么是「模板驱动」？没有模板可用吗？",
         "可用。未上传模板时，软件以通用规范（通用毕业论文格式）兜底；"
         "各校要求不一，上传本校模板，修正方更贴合。"),
        ("Q2 · 什么样的模板最好？一定要带批注吗？",
         "非必带，但带批注的学校官方模板效果最佳。软件定格式时，批注说明优先于样式定义："
         "批注有明确要求则遵批注，无批注则读样式定义，样式亦无则以通用规范兜底。"),
        ("Q3 · 试用版与正式版有何区别？卡片有哪些种类？",
         ("试用版：免费试用 %d 次，输出带水印的只读预览（不可直接编辑）。\n" % trial.TRIAL_LIMIT) +
         "正式版：无水印、可编辑、可保存。卡片共 4 种：\n"
         "· 永久卡 — 一次付费终身可用，激活后完全离线；\n"
         "· 周卡 / 月卡 — 期限内不限次数使用，到期后续费即可继续；\n"
         "· 次卡 — 按次计费，每修正一次扣一次次数（需联网扣次）。"),
        ("Q4 · 试用输出为只读，如何修改？",
         "试用文档设有只读保护；正式版输出可编辑文档，更为便捷，建议直接激活正式版。"),
        ("Q5 · 如何激活？激活后还需要联网吗？",
         "在激活窗口输入购买的激活码，验证通过即为正式版，右上角会显示您的卡片种类与剩余期限/次数。\n"
         "· 永久卡 — 激活后完全离线，无需再联网；\n"
         "· 周卡 / 月卡 — 激活后离线可用，到期前会联网校验一次；\n"
         "· 次卡 — 每次修正需联网扣一次次数，离线时暂不可用。"),
         ("Q6 · 论文安全吗？会上传吗？",
          "请放心。所有修正均在本机完成，论文不联网、不上传任何服务器，亦不收集论文内容。\n"
          "联网仅用于「激活校验」与「首次试用登记」，且只传输激活码 / 机器码，绝不含论文内容。\n"
          "永久卡 / 周卡 / 月卡激活后可断网使用；次卡仅在扣减次数时联网。"),
        ("Q7 · 学校模板特殊，或修正结果不尽如人意？",
         "欢迎关注公众号【%s】留言，告知贵校情况，我们协助处理。" % WECHAT_NAME),
        ("Q8 · 打开软件时 Windows 弹出“已保护你的电脑 / 已拦截”提示？",
         "本软件为未签名程序（省去每年数百元代码签名证书费用，让价格更亲民），Windows SmartScreen 会拦截提示，属正常现象，不代表软件有害。\n"
         "处理方式：在拦截框点击【详细信息】→ 再点击【仍要运行】即可正常打开；官网 reedskill.com 的“激活教程”页有图文演示。"),
        ("Q9 · 学生论文很多，一篇篇选太麻烦？",
          "用「选择文件夹」按钮：把学生论文统一放进一个文件夹，点此按钮，"
         "软件会自动收齐其中的全部 Word 文档（.docx / .doc / .wps），"
         "下方即时显示「已选择 N 篇论文」，随后可批量处理。\n"
         "注意：只扫描该文件夹内的文件，不会翻入其子文件夹，以免误选不相关的文档。"),
    ]
    for title, body in faqs:
        item(title, body)

    # —— 三、联系我们 ——
    section("三、联系我们")
    cl = albl(("使用疑问、模板咨询、购买事宜，均可通过以下方式联系：\n"
               "· 微信小程序：微信搜【%s】，可直接购买激活码，付款后自动发码（推荐）\n"
               "· 微信公众号：【%s】（ID：%s），关注后留言\n"
               "· 官网：reedskill.com（产品介绍 · 激活教程 · 购买方式）\n"
               "· 公众号设关键词自动回复（试用 / 激活 / 水印 / 报错），常见问题即时应答\n"
               "· 较复杂的问题，由人工回复\n"
               "· 联系邮箱：%s") % (MINIAPP_NAME, WECHAT_NAME, WECHAT_ID, ABOUT_MAIL),
              BODY, F_BODY)
    cl.pack(fill="x", padx=padx, pady=(8, 0))

    # 小程序码（购买主渠道）：图片缺失时自动跳过，不影响其余内容显示
    try:
        if os.path.isfile(MINIAPP_QRCODE):
            _mq = tk.PhotoImage(file=MINIAPP_QRCODE)
            _mq = _mq.subsample(max(1, round(_mq.width() / 130)))
            _mrow = tk.Frame(inner, bg=PAPER)
            _mrow.pack(pady=(10, 0))
            _mimg = tk.Label(_mrow, image=_mq, bg=PAPER)
            _mimg.image = _mq
            _mimg.pack()
            tk.Label(_mrow, text="微信扫一扫，或搜索【%s】小程序" % MINIAPP_NAME,
                     bg=PAPER, fg=MUTED, font=F_SMALL).pack(pady=(6, 0))
    except Exception:
        pass

    _site_label(inner, text="前往官网 reedskill.com →").pack(padx=padx, pady=(8, 0))

    tk.Label(inner, text="", bg=PAPER).pack(pady=(0, 18))


def main():
    # Windows 高分屏下让 Tk 字体清晰渲染（必须在创建 Tk 窗口前设置，否则整体发糊）
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    root = tk.Tk()
    _init_fonts(root)
    root.withdraw()
    if (not license.check_local_valid()) or license.is_revoked():
        # 未激活（或被后台撤销）时弹激活窗，但提供"先试用"入口（不进主界面则退出）
        r = show_activation(root)
        if r not in ("ok", "trial"):
            root.destroy()
            return
    root.deiconify()
    App(root)
    # 启动后台心跳：联网时若授权被撤销（退款），自动锁死软件
    if license.check_local_valid() and not license.is_revoked():
        license.start_heartbeat(product=license.DEFAULT_PRODUCT, on_revoked=_on_license_revoked)
    root.mainloop()


if __name__ == "__main__":
    main()
