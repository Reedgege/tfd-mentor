# -*- coding: utf-8 -*-
"""论文格式医生 · 导师版 · 无界面回归测试

覆盖所有引擎入口（这才是逻辑 bug 真正藏身处），不依赖 tkinter / 网络：
  - engine.run_check（无画像 / 带画像）
  - engine.run_fix_headings（modify=True 一键修正）
  - engine.run_annotate（modify=False 只批注不改原稿）
  - engine.run_reformat_refs（参考文献重排）
  - engine.run_build_profile（模板画像提取）
  - group_report.aggregate + xlsx_report.write_change_report_xlsx（汇总 Excel）
  - watermark.apply_watermark（水印 + 只读）
  - trial / license 激活边界（指向临时目录，不碰真实 ~/.tfd_mentor_license）

用法：python tfd-mentor/tests/test_headless_regression.py
退出码 0=全过，1=有失败。
"""
import os
import sys
import io
import json
import time
import shutil
import tempfile
import traceback
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))   # tfd-mentor/
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from tfd_app import engine
from tfd_app.core import group_report
from tfd_app.core import xlsx_report
from tfd_app import watermark
from tfd_app import trial
from tfd_app import license as lic

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# ---------------------------------------------------------------- 构造测试 docx
def _para(text, style=None):
    ppr = '<w:pPr><w:pStyle w:val="%s"/></w:pPr>' % style if style else ""
    return '<w:p>%s<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (ppr, text)

def _table():
    return (
        '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>'
        '<w:tblW w:w="0" w:type="auto"/></w:tblPr>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
        '<w:p><w:r><w:t>变量</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
        '<w:p><w:r><w:t>说明</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    )

def build_thesis(path):
    body = "".join([
        _para("摘要"),
        _para("本文研究毕业论文格式自动修正方法，按学校模板套用，验证引擎稳定性。"),
        _para("关键词"),
        _para("神经网络；模板驱动；格式修正"),
        _para("1 绪论"),
        _para("这是正文内容，用于验证正文区的格式套用与隔离。"),
        _para("1.1 研究背景"),
        _para("近年来，高校对毕业论文格式要求日益严格。"),
        _table(),
        _para("2 文献综述"),
        _para("已有研究从多个角度探讨了格式自动化。"),
        _para("参考文献"),
        _para("[1] 张三. 毕业论文格式规范研究[J]. 高等教育, 2020, 12(3): 45-50."),
        _para("[2] 李四. 自动化排版系统设计[D]. 北京: 某大学, 2021."),
        _para("致谢"),
        _para("感谢导师的悉心指导。"),
    ])
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="%s"><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'
           '</w:body></w:document>' % (W, body))
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="%s">'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/></w:style>'
              '<w:style w:type="paragraph" w:styleId="Heading1">'
              '<w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style>'
              '</w:styles>' % W)
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="word/styles.xml"/>'
            '</Relationships>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", styles)


def build_profile_json(path):
    """写一份最小但合法的画像（含 levels + headingStyles 规范式），供模板驱动路径使用。
    注意：headingStyles 值必须是对带 styleId 的字典（format_profile 产出格式），
    兼容 {lvl: "StyleId"} 旧式由 docxutils.normalize_heading_styles 在消费端兜底。
    """
    prof = {
        "source": "test_template.docx",
        "levels": {
            "abstract":  {"zh_font": "宋体", "en_font": "Times New Roman", "sz": "22", "bold": True, "align": "center"},
            "keywords":  {"zh_font": "宋体", "en_font": "Times New Roman", "sz": "22", "bold": True, "align": "left"},
            "ack":       {"zh_font": "黑体", "en_font": "Times New Roman", "sz": "22", "bold": True, "align": "center"},
        },
        "headingStyles": {
            "1": {"styleId": "Heading1", "name": "标题 1", "font": "黑体", "size": "小三", "outline_level": 1},
            "2": {"styleId": "Heading2", "name": "标题 2", "font": "黑体", "size": "四号", "outline_level": 2},
            "3": {"styleId": "Heading3", "name": "标题 3", "font": "黑体", "size": "小四", "outline_level": 3},
        },
        "body": {"styleId": "Normal", "font": "宋体", "size": "小四"},
        "page": {},
        "spec": {},
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(prof, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------- 测试框架
checks = []
def ck(name, ok, msg=""):
    checks.append((name, ok, msg))
    print(("  ✓ " if ok else "  ✗ ") + name + ((" | " + msg) if msg else ""))


def _is_valid_docx(p):
    if not os.path.isfile(p):
        return False
    try:
        with zipfile.ZipFile(p) as z:
            return "word/document.xml" in z.namelist()
    except Exception:
        return False


def run_all():
    tmp = tempfile.mkdtemp(prefix="tfd_regress_")
    try:
        src = os.path.join(tmp, "thesis.docx")
        profile = os.path.join(tmp, "profile.json")
        build_thesis(src)
        build_profile_json(profile)
        ck("构造测试 docx 合法", _is_valid_docx(src))

        # 1) 体检：无画像
        try:
            rep = engine.run_check(src)
            ck("run_check(无画像) 不崩且产出报告", isinstance(rep, str) and len(rep) > 0, "len=%d" % len(rep))
        except Exception:
            ck("run_check(无画像) 不崩且产出报告", False, traceback.format_exc().splitlines()[-1])

        # 2) 体检：带画像
        try:
            rep2 = engine.run_check(src, profile_path=profile)
            ck("run_check(带画像) 不崩且产出报告", isinstance(rep2, str) and len(rep2) > 0, "len=%d" % len(rep2))
        except Exception:
            ck("run_check(带画像) 不崩且产出报告", False, traceback.format_exc().splitlines()[-1])

        # 3) 一键修正（modify=True）
        dst_fix = os.path.join(tmp, "fixed.docx")
        try:
            md = engine.run_fix_headings(src, dst_fix, profile_path=profile, add_comments=False, modify=True)
            ck("run_fix_headings(修正) 不崩且产出 docx", _is_valid_docx(dst_fix), "md_len=%d" % len(md or ""))
        except Exception:
            ck("run_fix_headings(修正) 不崩且产出 docx", False, traceback.format_exc().splitlines()[-1])

        # 4) 只批注不改原稿（annotate）
        dst_ann = os.path.join(tmp, "annotated.docx")
        try:
            md2 = engine.run_annotate(src, dst_ann, profile_path=profile, author="导师")
            ck("run_annotate(只批注) 不崩且产出 docx", _is_valid_docx(dst_ann), "md_len=%d" % len(md2 or ""))
        except Exception:
            ck("run_annotate(只批注) 不崩且产出 docx", False, traceback.format_exc().splitlines()[-1])

        # 5) 参考文献重排
        dst_ref = os.path.join(tmp, "refs.docx")
        try:
            rep3 = engine.run_reformat_refs(src, dst_ref)
            ck("run_reformat_refs 不崩且产出 docx", _is_valid_docx(dst_ref), "rep_len=%d" % len(rep3 or ""))
        except Exception:
            ck("run_reformat_refs 不崩且产出 docx", False, traceback.format_exc().splitlines()[-1])

        # 6) 提取模板画像
        prof_out = os.path.join(tmp, "extracted.json")
        try:
            txt = engine.run_build_profile(src, prof_out)
            ok = os.path.isfile(prof_out) and len(txt) > 0
            ck("run_build_profile 不崩且产出 json", ok)
        except Exception:
            ck("run_build_profile 不崩且产出 json", False, traceback.format_exc().splitlines()[-1])

        # 7) 全组汇总：markdown( checker 格式 ) -> 抽问题 -> 归类 -> 聚合 -> Excel
        try:
            # 注意：extract_issues_from_markdown 只认 checker 固定格式 **[高/中/低/info]**
            md_sample = ("- 🔴 **[高]** 章节标题层级：第3章标题未按模板\n"
                        "- 🟠 **[中]** 参考文献格式不符 GB/T 7714\n"
                        "- 🔴 **[高]** 页边距偏大\n"
                        "- 🟠 **[中]** 章节标题层级：1.2 小节未套用")
            issues = group_report.extract_issues_from_markdown(md_sample)
            cats = [group_report.classify_issue(t) for t in issues]
            # aggregate 契约：[(论文标签, [类别,...]), ...]，返回 (summary, columns, rows)
            summary, columns, rows = group_report.aggregate(
                [("论文A.docx", cats), ("论文B.docx", cats)])
            ck("group_report 管线不崩", isinstance(rows, list) and len(rows) > 0,
               "rows=%d cols=%d" % (len(rows), len(columns)))
            # 直接用 aggregate 的产出（columns 为 (列名,宽度) 元组）驱动 xlsx，验证真实集成
            xlsx_path = os.path.join(tmp, "summary.xlsx")
            xlsx_report.write_change_report_xlsx(
                xlsx_path, "全组共性问题", summary, columns, rows)
            ck("xlsx_report 不崩且产出 xlsx", os.path.isfile(xlsx_path) and os.path.getsize(xlsx_path) > 0)
        except Exception:
            ck("group_report/xlsx_report 不崩", False, traceback.format_exc().splitlines()[-1])

        # 8) 水印 + 只读（作用在副本上，不破坏原文件）
        wm_src = os.path.join(tmp, "wm_src.docx")
        shutil.copy(src, wm_src)
        try:
            watermark.apply_watermark(wm_src)
            ok = _is_valid_docx(wm_src)
            ck("watermark.apply_watermark 不崩且产出合法 docx", ok)
        except Exception:
            ck("watermark.apply_watermark 不崩且产出合法 docx", False, traceback.format_exc().splitlines()[-1])

        # 9) trial / license 边界（指向临时目录，不影响真实授权）
        lic.LICENSE_DIR = tmp
        lic.LICENSE_FILE = os.path.join(tmp, "license.json")
        trial.TRIAL_FILE = os.path.join(tmp, "trial.json")
        # 隔离中台试用登记：单测不联网、不污染线上 trials 表（TRIAL_LIMIT=1）
        lic.server_trial_used = lambda mc: False
        lic.claim_server_trial = lambda mc: None
        trial._SERVER_USED.update(val=False, ts=1e18)   # 缓存置为「未用过」，避免联网
        try:
            assert not trial.is_licensed(), "测试环境不应已激活"
            ck("trial 初始剩余=1", trial.trials_left() == 1, "left=%d" % trial.trials_left())
            ck("trial 第1次扣减成功", trial.consume_trial() is True)
            ck("trial 用完第2次被拒", trial.consume_trial() is False)
            # 篡改锁定
            d = json.load(open(trial.TRIAL_FILE, encoding="utf-8"))
            d["used"] = 0
            json.dump(d, open(trial.TRIAL_FILE, "w", encoding="utf-8"))
            ck("trial 篡改后锁定(剩余0)", trial.trials_left() == 0)
            # 正式版绕开
            lic.save_local_license("TEST-CARD", lic.get_machine_code(), permanent=True)
            ck("激活后 is_licensed=True", trial.is_licensed() is True)
            ck("激活后扣减不消耗", trial.consume_trial() is True)
        except Exception:
            ck("trial/license 边界", False, traceback.format_exc().splitlines()[-1])

    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    all_ok = all(ok for _, ok, _ in checks)
    print("\n==== 导师版无界面回归: %s ====" % ("全部通过" if all_ok else "存在失败"))
    for name, ok, msg in checks:
        if not ok:
            print("  失败: %s (%s)" % (name, msg))
    return all_ok


if __name__ == "__main__":
    sys.exit(0 if run_all() else 1)
