# -*- coding: utf-8 -*-
"""
论文格式医生 · 中国版共享设计令牌（唯一来源）
==============================================

严格照施工包 ``03_SHARED/tokens/DESIGN_TOKENS.json`` + 各版 SVG 零件规格。
学生版 / 导师版共用同一套零件，只换调色板（edition 参数）。

配色铁律（AGENT_MASTER_INSTRUCTION.md）：
  - 宣纸米白底 + 学术蓝/墨青主色 + 少量印章红
  - 圆角：控件 8 / 卡片 12 / 弹窗 14 / 胶囊 999
  - 字号：唯一来源 = TYPE 表（照 _ui_ref/TYPE_SCALE.md；主视图须在 1280×720 整屏放下）
  - 无渐变、无玻璃拟态、无霓虹、无厚重阴影
  - 朱砂红只作身份印章 + 极克制点缀，不作主操作色
"""
from __future__ import annotations

import platform as _platform
import tkinter.font as tkfont

# ---------------------------------------------------------------------------
# 圆角 / 字号 / 间距（两版共用；字号唯一来源 = TYPE）
# ---------------------------------------------------------------------------
RADIUS = {"control": 8, "card": 12, "modal": 14, "pill": 999}

# 字号唯一权威表（_ui_ref/TYPE_SCALE.md 第 1 节，逐字照抄；单位 = 磅 pt）。
# 全仓库任何地方不得再出现字号字面量，一律 TYPE["token"] 取名 —— 新旧字号混排
# 正是「顶、挤、突兀」的成因。主视图要在 1280×720 整屏放下，字号按最小窗口倒推，
# 不照抄装修包规格书的浏览器 px（海外版同一条纪律：F_BRAND16/F_BODY9/F_FOOT8）。
# 层级比例（TYPE_SCALE.md 第 3 节）：20/14=1.43、14/10=1.4、20/10=2.0、12/10=1.2。
TYPE = {
    "page_title": 20,     # 顶部大标题「论文格式医生」（全界面最大字号，仅此一处）
    "subtitle": 10,       # 副标题「让论文格式更简单」
    "section_title": 14,  # 区块标题「文件选择」「处理步骤」
    "card_title": 12,     # 步骤名 / 卡片标题 / 弹窗小节标题
    "body": 10,           # 正文、说明、字段值、单选行
    "caption": 9,         # 次要说明小字、行内帮助
    "btn": 10,            # 主按钮
    "btn_small": 9,       # 次按钮（描边）
    "step_num": 9,        # 步骤圆点里的数字
    "seal": 12,           # 朱砂竖章（竖排「导师」）
    "footer": 9,          # 页脚 / 状态栏
    "mono": 9,            # 机器码 / 离线码（Consolas）
}

# 间距 / 固定像素高度：字号缩小后同步收紧，避免「字小了、块还是那么大」的空洞感。
SPACING = {
    "hair": 2, "xs": 4, "tight": 6, "sm": 8, "md": 12, "lg": 16, "xl": 24, "xxl": 32,
    "page_mx": 30,       # 页面左右留白
    "card_pad": 12,      # 卡片内边距（字号缩小后同步收紧，不留空洞）
    "card_gap": 18,      # 两卡之间的一档留白
    "head_h": 88,        # 顶部品牌区高度（品牌组高 81 + 上下各约 3px）
    "foot_h": 58,        # 页脚高度（细线 + 两行文字）
    "status_h": 20,      # 状态栏高度
    "field_h": 28,       # 输入行 / 次按钮高度
    "btn_h": 34,         # 主按钮 / 弹窗按钮高度
    "btn_h_sm": 24,      # 小按钮（危险按钮 / 徽章内按钮）高度
    "badge_h": 22,       # 胶囊徽标高度
    "dropzone_h": 68,    # 上传区固定高（12pt 主文案 + 9pt 副文案的最小不裁切高度）
    "step_line": 40,     # 步进器连接线长度
    "mark_h": 36,        # 品牌 mark 高度（TYPE_SCALE.md 第 2 节：34～38px）
    "mark_w": 46,        # 品牌 mark 宽度
}

# ---------------------------------------------------------------------------
# 两版调色板（照 DESIGN_TOKENS.json 的 student / advisor 段）
# ---------------------------------------------------------------------------
_PALETTES = {
    "student": {
        "bg": "#F7F3EA",          # 宣纸米白
        "surface": "#FFFDF9",     # 卡片面
        "primary": "#286D9F",     # 学术蓝（主操作色）
        "navy": "#193E5D",        # 标题深蓝
        "seal": "#B54D43",        # 印章红
        "muted": "#6F8493",       # 次要文字
        "border": "#D7E0E5",      # 卡片 / 控件边框
    },
    "advisor": {
        "bg": "#F5EFE5",          # 更暖的米白
        "surface": "#FFFDF9",
        "primary": "#355F73",     # 墨青（主操作色）
        "navy": "#3F3632",        # 暖褐黑
        "seal": "#B54D43",        # 印章红（同）
        "muted": "#776E68",
        "border": "#D9D0C5",
    },
}

# 由主色派生（主色加深作 hover；主色极浅作徽标底）
def _derive(edition, base):
    is_student = (edition == "student")
    return {
        "primary_hover": "#1F5A8A" if is_student else "#2C4F60",
        "primary_soft": "#EAF3FB" if is_student else "#EAF0F2",
        "secondary_border": "#BFD0DA",     # 次按钮 / 选择行边框（SVG 实测）
        "dropzone_fill": "#F7FBFE",
        "dropzone_border": "#BBD2E2",
        "info_fill": "#F0F6FA" if is_student else "#F1F4F2",
        "info_border": "#D5E3EC" if is_student else "#DCD3C8",
        "select_border": "#D0DCE4",
        "select_ph": "#506B7C",
        "stepper_line": "#D8E2EA",
        "step_pending_fill": "#FFFDF9",
        "step_pending_border": "#BFD0DA",
        "stepper_label_active": "#284E68" if is_student else "#3F3632",
        "stepper_label_pending": "#6F8493",
        # 状态三色（03_SHARED/components 公共，两版一致）
        "success_fill": "#F1F7F3", "success_border": "#D0E2D8",
        "success_dot": "#3E7B67", "success_text": "#3E6C5D",
        "processing_fill": "#F1F6FA", "processing_border": "#D2E0EA",
        "processing_dot": "#286D9F", "processing_text": "#355F73",
        "error_fill": "#FFF5F3", "error_border": "#E5C8C2",
        "error_dot": "#A94A43", "error_text": "#98453F",
        # 正式版徽标（green-grey，区别于试用蓝）
        "formal_fill": "#F0F2EF", "formal_text": "#4D6C61",
        # 危险按钮（danger-button.svg：浅朱砂底 + 朱砂描边 + 朱砂字）
        "danger_fill": "#FFF8F6", "danger_border": "#D8B0AA",
        "danger_text": "#A94A43", "danger_hover": "#FBEDEA",
        # 浅色小徽章（视觉稿「可选」）
        "soft_fill": "#F0F6FA" if is_student else "#F1F4F2",
        "soft_border": "#D5E3EC" if is_student else "#DCD3C8",
        "soft_text": "#50708A" if is_student else "#6B6259",
        # 单选（radio-on.svg / radio-off.svg）
        "radio_on": "#286D9F" if is_student else "#355F73",
        "radio_off_border": "#AFC0CB" if is_student else "#BFB3A6",
        # 页脚细线 / 文案色（footer.svg）
        "footer_line": "#D8E2EA",
        "footer_text": "#6F8493",
        "footer_text_soft": "#8A979E",
        # 弹窗壳
        "modal_fill": "#FFFDF9", "modal_border": "#D8E2EA",
        "modal_title": "#193E5D", "modal_sub": "#6F8493",
        # 卡片标题（区块标题用 navy）
        "card_title": base["navy"],
        # 禁用手写色
        "disabled_fill": "#ECE9E2" if is_student else "#E7E1D6",
        "disabled_text": "#9AA1A5",
        "disabled_border": "#ECE9E2" if is_student else "#E7E1D6",
    }


class Theme:
    """edition 调色板 + 派生色 + 字体解析，集中一处避免散落硬编码。"""

    def __init__(self, edition):
        self.edition = edition
        base = _PALETTES[edition]
        self.__dict__.update(base)
        self.__dict__.update(_derive(edition, base))
        self.edition_label = "学生版" if edition == "student" else "导师版"

    # ---- 字体（中文标题走宋体/思源宋体；正文走雅黑/思源黑体；拉丁数字走 Georgia）----
    @staticmethod
    def _resolve(fam):
        fams = tkfont.families()
        if fam in fams:
            return fam
        mapping = {
            "SimSun": {"Darwin": "Songti SC", "Linux": "Noto Serif CJK SC"},
            "Microsoft YaHei": {"Darwin": "PingFang SC", "Linux": "Noto Sans CJK SC"},
            "Consolas": {"Darwin": "Menlo", "Linux": "DejaVu Sans Mono"},
            "Georgia": {},
        }
        cand = mapping.get(fam, {}).get(_platform.system())
        if cand and cand in fams:
            return cand
        return fam

    def serif(self, size, bold=False):
        return (self._resolve("SimSun"), size, "bold" if bold else "normal")

    def sans(self, size, bold=False):
        return (self._resolve("Microsoft YaHei"), size, "bold" if bold else "normal")

    def latin(self, size, bold=False):
        return ("Georgia", size, "bold" if bold else "normal")

    def mono(self, size, bold=False):
        return (self._resolve("Consolas"), size, "bold" if bold else "normal")


_THEME_CACHE = {}


def get_theme(edition="student"):
    if edition not in _THEME_CACHE:
        _THEME_CACHE[edition] = Theme(edition)
    return _THEME_CACHE[edition]
