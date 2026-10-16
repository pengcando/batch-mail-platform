# -*- coding: utf-8 -*-
"""
xlsx 样式搬运工具。

背景：生成的《公司名&xxx对账单-年月.xlsx》需要"长得跟模板一模一样"，
而 openpyxl 不支持跨工作簿直接复制 worksheet，所以这里手工把
**单元格样式 / 列宽 / 行高 / 冻结窗格** 从模板搬运到新工作簿。

要点：
- 样式对象必须 copy()，否则多个单元格会共享同一个对象，改一个动一片。
- 只搬「格式」，不搬「公式」。原模板的 `对账单` 页全是 `='账单数据-原始数据'!X` 公式，
  线上化后由引擎直接算成字面值写入，因此新文件里不含任何外部引用（不会出现 #REF!）。
"""
from copy import copy


def apply_style(style, dst_cell):
    """把快照好的样式对象套到目标单元格上。"""
    dst_cell.font = copy(style["font"])
    dst_cell.border = copy(style["border"])
    dst_cell.fill = copy(style["fill"])
    dst_cell.alignment = copy(style["alignment"])
    dst_cell.protection = copy(style["protection"])
    dst_cell.number_format = style["number_format"]


def snapshot_style(cell):
    """把单元格的样式冻结成一份可复用的快照。"""
    return {
        "font": copy(cell.font),
        "border": copy(cell.border),
        "fill": copy(cell.fill),
        "alignment": copy(cell.alignment),
        "protection": copy(cell.protection),
        "number_format": cell.number_format,
    }


class SheetTemplate:
    """一张 sheet 的"版式模板"：表头行 + 数据行样式 + 列宽行高。"""

    def __init__(self, ws, ncols, header_row=1, data_row=2):
        self.ncols = ncols
        self.header_row = header_row
        self.data_row = data_row

        self.title = ws.title
        self.header_values = [
            ws.cell(row=header_row, column=c).value for c in range(1, ncols + 1)
        ]
        self.header_styles = [
            snapshot_style(ws.cell(row=header_row, column=c)) for c in range(1, ncols + 1)
        ]
        self.data_styles = [
            snapshot_style(ws.cell(row=data_row, column=c)) for c in range(1, ncols + 1)
        ]

        # 注意：data_row 行可能不存在（空表），做一次兜底
        if ws.max_row < data_row:
            self.data_styles = self.header_styles

        self.header_height = _dim_height(ws.row_dimensions, header_row)
        self.data_height = _dim_height(ws.row_dimensions, data_row)

        self.column_widths = {}
        for key, dim in ws.column_dimensions.items():
            if dim.width:
                self.column_widths[key] = {
                    "width": dim.width,
                    "hidden": dim.hidden,
                    "bestFit": dim.bestFit,
                    "outlineLevel": dim.outlineLevel,
                }
        self.freeze_panes = ws.freeze_panes


def _dim_height(container, idx):
    dim = container.get(idx)
    return dim.height if dim is not None else None


def build_sheet(wb, template, rows, title=None):
    """按模板生成一张新 sheet：表头取模板第 1 行，数据样式取模板第 2 行。"""
    ws = wb.create_sheet(title or template.title)

    for c in range(1, template.ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.value = template.header_values[c - 1]
        apply_style(template.header_styles[c - 1], cell)

    for r_offset, row_values in enumerate(rows):
        excel_row = 2 + r_offset
        for c in range(1, template.ncols + 1):
            cell = ws.cell(row=excel_row, column=c)
            value = row_values[c - 1] if c - 1 < len(row_values) else None
            if value is not None:
                cell.value = value
            apply_style(template.data_styles[c - 1], cell)

    # 列宽
    for key, info in template.column_widths.items():
        dim = ws.column_dimensions[key]
        dim.width = info["width"]
        dim.hidden = info["hidden"]
        dim.bestFit = info["bestFit"]
        dim.outlineLevel = info["outlineLevel"]
    # 行高
    if template.header_height:
        ws.row_dimensions[1].height = template.header_height
    if template.data_height:
        for excel_row in range(2, 2 + len(rows)):
            ws.row_dimensions[excel_row].height = template.data_height

    # 冻结窗格（模板有就跟着冻结，没有就保持不动）
    if template.freeze_panes:
        ws.freeze_panes = template.freeze_panes

    return ws


def set_sheet_number_format(ws, fmt, max_row=None, max_col=None):
    """把整张表（含表头）的单元格格式统一改成 fmt —— 对应 流程文档的 B15「整表设为数值格式」。"""
    max_row = max_row or ws.max_row
    max_col = max_col or ws.max_column
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            cell.number_format = fmt


# ---------------------------------------------------------------------------
# M 表的行重排（对应流程文档 A11「A 列降序」）
# ---------------------------------------------------------------------------

def reorder_rows_by_key(ws, key_index, reverse=True, first_row=2, last_row=None, ncols=None):
    """
    按指定列把数据行整体重排（默认降序），返回重排后每一行的 Excel 绝对行号映射。

    实现方式：把每行的「值 + 样式」一起读出来，清空后按新顺序写回。
    这样即使各行样式不同也不会错位。
    """
    last_row = last_row or ws.max_row
    ncols = ncols or ws.max_column
    if last_row < first_row:
        return []

    records = []
    for r in range(first_row, last_row + 1):
        values = [ws.cell(row=r, column=c).value for c in range(1, ncols + 1)]
        styles = [snapshot_style(ws.cell(row=r, column=c)) for c in range(1, ncols + 1)]
        key = values[key_index - 1]
        records.append({"values": values, "styles": styles, "key": "" if key is None else str(key)})
        # 先清空，避免旧值残留
        for c in range(1, ncols + 1):
            ws.cell(row=r, column=c).value = None

    records.sort(key=lambda x: x["key"], reverse=reverse)

    for offset, rec in enumerate(records):
        r = first_row + offset
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.value = rec["values"][c - 1]
            apply_style(rec["styles"][c - 1], cell)

    return records
