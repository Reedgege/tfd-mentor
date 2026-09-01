# -*- coding: utf-8 -*-
"""
标准库手写 xlsx（零第三方依赖）：把修改清单导出为 Excel。

导师版用 Excel 呈现修改明细，方便导师/学生横向核对。
仅依赖 zipfile + xml，不引入 openpyxl，保证打包后零第三方依赖。

用法：
    from xlsx_report import write_change_report_xlsx
    write_change_report_xlsx(path, "论文格式修改明细", summary, columns, rows)
    # columns = [("列名", 宽度px), ...]
    # rows    = [[cell, ...], ...]    与 columns 等长
    # summary = [str, ...]            顶部摘要行
"""
import os
import zipfile


def _esc(s):
    """XML 文本转义（含单/双引号，避免属性断句）。"""
    return (str(s)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&apos;"))


def _col_letter(idx):
    """0-based 列号 → Excel 列字母（A, B, ..., Z, AA, ...）。"""
    letters = ""
    n = idx
    while True:
        letters = chr(ord("A") + (n % 26)) + letters
        n = n // 26 - 1
        if n < 0:
            break
    return letters


def _disp_w(s):
    """估算单元格显示宽度：中文/全角字符按 2 计，其余按 1（Excel 列宽单位约 1 拉丁字符）。"""
    n = 0
    for ch in str(s):
        n += 2 if ord(ch) > 0x2E80 else 1
    return n


# 样式表（cellXfs 索引即调用方使用的 s 值）：
# 0 默认 / 1 标题(bold大) / 2 表头(白字朱砂底+居中) /
# 3 数据格(细边框+左对齐+自动换行) / 4 低置信度(红字+边框) / 5 摘要行(bold+换行)
_STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<fonts count="5">'
    '<font><sz val="11"/><name val="\u5b8b\u4f53"/></font>'
    '<font><b/><sz val="14"/><name val="\u5b8b\u4f53"/></font>'
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="\u5b8b\u4f53"/></font>'
    '<font><sz val="11"/><color rgb="FFFF0000"/><name val="\u5b8b\u4f53"/></font>'
    '<font><b/><sz val="11"/><name val="\u5b8b\u4f53"/></font>'
    '</fonts>'
    '<fills count="3">'
    '<fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FF9A3B34"/><bgColor indexed="64"/></patternFill></fill>'
    '</fills>'
    '<borders count="2">'
    '<border><left/><right/><top/><bottom/><diagonal/></border>'
    '<border>'
    '<left style="thin"><color rgb="FFD8CDB8"/></left>'
    '<right style="thin"><color rgb="FFD8CDB8"/></right>'
    '<top style="thin"><color rgb="FFD8CDB8"/></top>'
    '<bottom style="thin"><color rgb="FFD8CDB8"/></bottom>'
    '<diagonal/>'
    '</border>'
    '</borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="6">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>'
    '<xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>'
    '<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>'
    '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '</styleSheet>'
)


def write_change_report_xlsx(path, title, summary, columns, rows):
    """写出 Excel 修改明细（.xlsx）。

    path:    输出文件名；title: 表标题；summary: 摘要行列表；
    columns: [(列名, 宽度px), ...]；rows: [[单元格, ...], ...]。
    表头加粗朱砂底、冻结表头、置信度"低"红字提示。
    """
    n_cols = len(columns)
    col_letters = [_col_letter(i) for i in range(n_cols)]
    last_col = col_letters[-1]

    # 列宽：按每列真实内容自适应，并封顶（避免"内容摘要"等列被写死成 300+ 字符宽）。
    # 取表头与各单元格中最长显示宽度，限制在区间 [_MIN_W, _MAX_W] 内；超出部分靠自动换行 + 行高自适应解决。
    _MIN_W, _MAX_W = 8, 46
    cols_xml = []
    for i, (name, _w) in enumerate(columns):
        m = _disp_w(name)
        for row in rows:
            if i < len(row) and row[i] is not None:
                m = max(m, _disp_w(row[i]))
        cw = max(_MIN_W, min(_MAX_W, m))
        cols_xml.append('<col min="%d" max="%d" width="%d" customWidth="1"/>' % (i + 1, i + 1, cw))
    cols_xml = "<cols>" + "".join(cols_xml) + "</cols>"

    # 行布局
    title_row = 1
    blank1 = 2
    summary_start = 3
    if summary:
        summary_end = 2 + len(summary)
        blank2 = summary_end + 1
    else:
        blank2 = summary_start
    header_row = blank2 + 1
    data_start = header_row + 1

    sheet_rows = []

    # 标题（合并整行）
    sheet_rows.append(
        '<row r="%d"><c r="A%d" s="1" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c></row>'
        % (title_row, title_row, _esc(title)))

    # 摘要行（合并整行）
    if summary:
        for idx, line in enumerate(summary):
            r = summary_start + idx
            sheet_rows.append(
                '<row r="%d"><c r="A%d" s="5" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c></row>'
                % (r, r, _esc(line)))

    # 表头
    hdr = []
    for i, (name, w) in enumerate(columns):
        hdr.append('<c r="%s%d" s="2" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                   % (col_letters[i], header_row, _esc(name)))
    sheet_rows.append('<row r="%d">%s</row>' % (header_row, "".join(hdr)))

    # 数据行
    for ri, row in enumerate(rows):
        r = data_start + ri
        cells = []
        for ci, val in enumerate(row):
            col = col_letters[ci]
            s = "3"
            if columns[ci][0] == "\u7f6e\u4fe1\u5ea6" and str(val).startswith("\u4f4e"):
                s = "4"  # 置信度列且为"低…" → 红字
            cells.append('<c r="%s%d" s="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                         % (col, r, s, _esc(val)))
        sheet_rows.append('<row r="%d">%s</row>' % (r, "".join(cells)))

    # 合并区域：标题 + 各摘要行跨所有列
    merges = ['<mergeCell ref="A%d:%s%d"/>' % (title_row, last_col, title_row)]
    if summary:
        for idx in range(len(summary)):
            r = summary_start + idx
            merges.append('<mergeCell ref="A%d:%s%d"/>' % (r, last_col, r))
    merges_xml = '<mergeCells count="%d">%s</mergeCells>' % (len(merges), "".join(merges))

    # 冻结表头（及其上方摘要）
    freeze = ('<sheetViews><sheetView workbookViewId="0">'
              '<pane ySplit="%d" topLeftCell="A%d" activePane="bottomLeft" state="frozen"/>'
              '</sheetView></sheetViews>' % (header_row, data_start))

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + freeze
        + '<sheetFormatPr defaultRowHeight="15"/>'
        + cols_xml
        + '<sheetData>' + "".join(sheet_rows) + '</sheetData>'
        + merges_xml
        + '</worksheet>'
    )

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>'
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="\u4fee\u6539\u660e\u7ec6" sheetId="1" r:id="rId1"/></sheets>'
        '</workbook>'
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>'
    )

    parent = os.path.dirname(os.path.abspath(path))
    if parent:
        os.makedirs(parent, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook_xml)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        z.writestr("xl/styles.xml", _STYLES_XML)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
    return path


if __name__ == "__main__":
    # 冒烟自测：生成一份示例 xlsx（用浏览器/Excel 打开验证）
    _summary = ["修改模式：通用 Heading 样式（通用 Heading 样式）", "共修改 2 处（2 条记录：标题 1 处 / 正文 1 处）"]
    _cols = [("\u5e8f\u53f7", 700), ("\u7c7b\u578b", 1000), ("\u4fee\u6539\u524d", 2100),
             ("\u4fee\u6539\u540e", 2100), ("\u5185\u5bb9\u6458\u8981", 2260), ("\u7f6e\u4fe1\u5ea6", 900)]
    _rows = [["1", "\u4e00\u7ea7", "1 \u7eea\u8bba", "\u7eea\u8bba", "\u6807\u9898\u7f16\u53f7\u89c4\u8303\u5316", "\u9ad8"],
             ["2", "\u6b63\u6587", "\u5c0f\u4e8c\u9ed1", "\u5b8b\u4f53", "\u5b57\u4f53\u7edf\u4e00", "\u4f4e\uff08\u5efa\u8bae\u590d\u6838\uff09"]]
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "smoke_report.xlsx")
    write_change_report_xlsx(out, "\u8bba\u6587\u683c\u5f0f\u4fee\u6539\u660e\u7ec6\uff08\u6807\u9898\u4e0e\u6b63\u6587\uff09",
                            _summary, _cols, _rows)
    print("示例 xlsx 已生成:", os.path.abspath(out))
