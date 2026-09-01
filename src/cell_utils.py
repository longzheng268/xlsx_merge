"""
单元格操作模块 - 负责单元格值、公式、样式的深拷贝

职责单一：只做单元格级别的读写操作，不涉及业务逻辑。
依赖：openpyxl（外部库）
"""

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter


def copy_cell_style(src_cell: openpyxl.cell.Cell, target_cell: openpyxl.cell.Cell) -> None:
    """
    深拷贝单元格样式（字体、边框、填充、对齐、数字格式、保护）。
    不复制值和公式，仅复制视觉格式。
    """
    if src_cell.has_style:
        target_cell.font = src_cell.font.copy()
        target_cell.border = src_cell.border.copy()
        target_cell.fill = src_cell.fill.copy()
        target_cell.number_format = src_cell.number_format
        target_cell.protection = src_cell.protection.copy()
        target_cell.alignment = src_cell.alignment.copy()


def copy_cell_value(src_cell: openpyxl.cell.Cell, target_cell: openpyxl.cell.Cell) -> None:
    """
    复制单元格的值或公式。
    若值以 '=' 开头则视为公式，原样保留；否则复制普通值。
    """
    target_cell.value = src_cell.value


def copy_cell(src_cell: openpyxl.cell.Cell, target_cell: openpyxl.cell.Cell) -> None:
    """完整复制单元格：值/公式 + 样式 + 批注"""
    copy_cell_value(src_cell, target_cell)
    copy_cell_style(src_cell, target_cell)
    if src_cell.comment is not None:
        try:
            target_cell.comment = src_cell.comment.copy()
        except Exception:
            target_cell.comment = src_cell.comment


def copy_row(src_ws: Worksheet, src_row: int,
             target_ws: Worksheet, target_row: int,
             max_col: int, row_offset: int = 0,
             adjust_formulas: bool = False) -> None:
    """
    复制整行数据（含样式）从源Sheet到目标Sheet。
    """
    src_height = src_ws.row_dimensions[src_row].height
    if src_height is not None:
        target_ws.row_dimensions[target_row].height = src_height

    for col in range(1, max_col + 1):
        src_cell = src_ws.cell(row=src_row, column=col)
        tgt_cell = target_ws.cell(row=target_row, column=col)
        copy_cell_value(src_cell, tgt_cell)
        copy_cell_style(src_cell, tgt_cell)
        if src_cell.comment is not None:
            try:
                tgt_cell.comment = src_cell.comment.copy()
            except Exception:
                tgt_cell.comment = src_cell.comment

        if adjust_formulas and row_offset != 0:
            _adjust_formula_in_cell(tgt_cell, row_offset)


def copy_column_widths(src_ws: Worksheet, target_ws: Worksheet, max_col: int) -> None:
    """
    将源Sheet的列宽复制到目标Sheet。
    仅在第一次写入表头时调用。
    """
    for col in range(1, max_col + 1):
        col_letter = get_column_letter(col)
        src_width = src_ws.column_dimensions[col_letter].width
        if src_width is not None:
            target_ws.column_dimensions[col_letter].width = src_width


def copy_merged_cells(src_ws: Worksheet, target_ws: Worksheet,
                      src_row_start: int, src_row_end: int,
                      target_row_start: int, row_offset: int) -> None:
    """
    复制源Sheet中指定行范围内的合并单元格信息到目标Sheet。
    仅复制完全在 [src_row_start, src_row_end] 范围内的合并区域。
    """
    from openpyxl.worksheet.cell_range import CellRange

    for merged_range in src_ws.merged_cells.ranges:
        mr_min = merged_range.min_row
        mr_max = merged_range.max_row
        # 仅复制完全在范围内的合并区域
        if mr_min >= src_row_start and mr_max <= src_row_end:
            new_min_row = mr_min + row_offset
            new_max_row = mr_max + row_offset
            new_range = CellRange(
                min_col=merged_range.min_col,
                max_col=merged_range.max_col,
                min_row=new_min_row,
                max_row=new_max_row,
            )
            target_ws.merge_cells(str(new_range))


def _adjust_formula_in_cell(cell: openpyxl.cell.Cell, row_offset: int) -> None:
    """
    对单个单元格内的公式进行行偏移修正。
    使用 openpyxl.formula.translate.Translator 实现精准偏移。

    注意：对于引用表头行的相对引用（如 H5），翻译后可能指向错误行。
    建议仅在确认公式不引用固定行时启用此功能。
    翻译失败或结果无效时，保留原始公式并输出警告。
    """
    if cell.value is None or not isinstance(cell.value, str) or not cell.value.startswith('='):
        return

    original = cell.value
    try:
        from openpyxl.formula.translate import Translator
        translator = Translator(original, cell.coordinate)
        translated = translator.translate(row_offset)

        # 验证翻译结果：必须是非空字符串且以 = 开头
        if translated and isinstance(translated, str) and translated.startswith('='):
            cell.value = translated
        else:
            print(f"  [公式警告] {cell.coordinate}: 翻译结果无效，保留原公式")
            cell.value = original
    except Exception as e:
        # 翻译失败时保留原始公式，避免数据丢失
        print(f"  [公式警告] {cell.coordinate}: 翻译异常({type(e).__name__})，保留原公式")
        cell.value = original


def is_effective_row(ws: Worksheet, row: int, max_col: int) -> bool:
    """判断一行是否包含可输出的真实数据：A列单独有值不算"""
    for col in range(2, max_col + 1):
        val = ws.cell(row=row, column=col).value
        if val is not None and str(val).strip() != '':
            return True
    return False


def should_copy_row(ws: Worksheet, row: int, max_col: int) -> bool:
    """判断一行是否应复制到输出：至少要有B列及之后的有效内容"""
    return is_effective_row(ws, row, max_col)


def is_row_empty(ws: Worksheet, row: int, max_col: int) -> bool:
    """判断指定行是否为空行（A列单独有值也视为空行）"""
    return not should_copy_row(ws, row, max_col)