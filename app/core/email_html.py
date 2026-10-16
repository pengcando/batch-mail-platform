# -*- coding: utf-8 -*-
"""
邮件正文渲染 —— 对应流程文档 §4.1。

原流程是「截图 开票信息!A1:I10 → Base64 内嵌图片」，网页端改为
「读取单元格值 → 渲染 HTML」，纯文本、体积小、全客户端可见。

正文结构（自上而下）：
    1. 前导空行（3 个 &nbsp; 空段）
    2. 开票信息区：A1 长文本按换行拆成多个 <p>；A5 提示语；A6:C10 两列表格
    3. 落款区：4 空格缩进
"""

PRE_BLANK_LINES = 3          # 原文第 1 段的空行数量
SIGNATURE_INDENT = "&nbsp;" * 4

FOOTER_LINES = [
    "商户对账",
    "",
    "示例公司 | example.com",
    "Email: sender@example.com",
]


def fmt_amount(value):
    """金额统一两位小数（业务确认）。非数字原样返回。"""
    if value is None or value == "":
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)):
        return f"{value:,.2f}"
    text = str(value).strip()
    try:
        return f"{float(text.replace(',', '')):,.2f}"
    except ValueError:
        return text


def _esc(text):
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def render_invoice_html(invoice):
    """
    invoice 结构（见 engine.build_invoice_context）：
        {
          "intro": "A1 长文本",
          "tip":   "A5 提示语",
          "rows":  [("商户品牌名称：", "爱点饭", False), ..., ("服务费开票金额：", 0.0, True)]
        }
    第三个元素 is_amount 决定是否按两位小数渲染。
    """
    parts = []

    for _ in range(PRE_BLANK_LINES):
        parts.append("<p>&nbsp;</p>")

    intro = (invoice.get("intro") or "").replace("\r\n", "\n").replace("\r", "\n")
    for line in intro.split("\n"):
        parts.append(f"<p>{_esc(line) if line.strip() else '&nbsp;'}</p>")

    tip = invoice.get("tip") or ""
    if tip:
        parts.append(f"<p>{_esc(tip)}</p>")

    table_rows = []
    for label, value, is_amount in invoice.get("rows", []):
        if is_amount:
            shown = fmt_amount(value)
        else:
            shown = "" if value is None else str(value)
        table_rows.append(
            "<tr>"
            f'<td style="padding:3px 24px 3px 0;white-space:nowrap;">{_esc(label)}</td>'
            f'<td style="padding:3px 0;">{_esc(shown)}</td>'
            "</tr>"
        )
    if table_rows:
        parts.append(
            '<table cellpadding="0" cellspacing="0" border="0" '
            'style="border-collapse:collapse;font-family:inherit;">'
            + "".join(table_rows)
            + "</table>"
        )

    parts.append("<p>&nbsp;</p>")
    for line in FOOTER_LINES:
        if line:
            parts.append(f"<p>{SIGNATURE_INDENT}{_esc(line)}</p>")
        else:
            parts.append("<p>&nbsp;</p>")
    parts.append("<p>&nbsp;</p>")

    return "\n".join(parts)


def render_invoice_text(invoice):
    """纯文本兜底版本（部分老客户端只认 text/plain）。"""
    lines = [""] * PRE_BLANK_LINES
    intro = (invoice.get("intro") or "").replace("\r\n", "\n")
    lines.extend(intro.split("\n"))
    if invoice.get("tip"):
        lines.append(invoice["tip"])
    for label, value, is_amount in invoice.get("rows", []):
        shown = fmt_amount(value) if is_amount else ("" if value is None else str(value))
        lines.append(f"{label}\t{shown}")
    lines.append("")
    for line in FOOTER_LINES:
        lines.append(("    " + line) if line else "")
    return "\n".join(lines)


def wrap_html(body_fragment):
    """包一层最小 HTML 壳，保证中文与字体在各客户端一致。"""
    return (
        '<div style="font-family:\'Microsoft YaHei\',\'微软雅黑\',Arial,sans-serif;'
        'font-size:10.5pt;color:#222;line-height:1.6;">'
        + body_fragment
        + "</div>"
    )
