"""
核心合并引擎模块 - 编排整个合并流程

职责：协调 SheetAnalyzer、Validator、CellUtils 完成合并。
本身不包含具体的单元格操作或校验逻辑。
"""

import os
import openpyxl
from typing import List, Optional

from config import MergeConfig
from sheet_analyzer import analyze_sheet, describe_region, SheetRegion, SIGNATURE_KEYWORDS, LEGEND_KEYWORDS
from validator import validate_sheet, validate_all_sheets, ValidationResult
from cell_utils import (
    copy_row, copy_column_widths, copy_merged_cells, is_row_empty, should_copy_row, is_effective_row,
    detect_attendance_day_columns, ATTENDANCE_COL_START,
)


MERGE_OUTPUT_MARKERS = ("成功合并", "合并状态", "合并结果位置")


def _is_merge_output(wb: openpyxl.Workbook) -> bool:
    """判断一个工作簿是否是本工具的历史合并产物（含“报告/总表”汇总页）。

    若首个 Sheet 的左上角出现“成功合并/合并状态/合并结果位置”等标记，则视为
    合并产物，避免把上一次的输出文件当成输入再次合并。
    """
    if not wb.sheetnames:
        return False
    try:
        ws = wb[wb.sheetnames[0]]
    except Exception:
        return False
    for r in range(1, min(ws.max_row, 10) + 1):
        for c in range(1, min(ws.max_column, 10) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                for m in MERGE_OUTPUT_MARKERS:
                    if m in v:
                        return True
    return False


class SheetUnit:
    def __init__(self, file_name: str, sheet_name: str,
                 wb: openpyxl.Workbook, ws: openpyxl.worksheet.worksheet.Worksheet):
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.wb = wb
        self.ws = ws
        self.region: Optional[SheetRegion] = None

    def analyze(self, config: MergeConfig) -> None:
        self.region = analyze_sheet(
            self.ws,
            header_rows=config.header_rows,
            footer_rows=config.footer_rows,
            trim_trailing=config.trim_trailing_empty,
        )

    @property
    def is_empty(self) -> bool:
        return self.region is None or self.region.real_max_row == 0


class MergeEngine:
    def __init__(self, config: MergeConfig):
        self.config = config
        self.sheet_units: List[SheetUnit] = []
        self.validation_results: List[ValidationResult] = []
        self._log_lines: List[str] = []
        self._col_widths = self._collect_column_widths()

    def scan_files(self) -> List[str]:
        input_dir = self.config.input_dir
        output_name = os.path.basename(self.config.output_file)
        files = []
        for f in os.listdir(input_dir):
            if not f.endswith('.xlsx') or f.startswith('~$'):
                continue
            if f == output_name:
                self._log(f"[跳过] 输出文件本身不作为输入: {f}")
                continue
            files.append(f)
        files.sort()
        if not files:
            self._log("未找到可处理的 .xlsx 文件")
        else:
            self._log(f"找到 {len(files)} 个待合并文件:")
            for f in files:
                self._log(f"  - {f}")
        return files

    def load_sheet_units(self) -> None:
        files = self.scan_files()
        self.sheet_units = []

        for fname in files:
            fpath = os.path.join(self.config.input_dir, fname)
            try:
                wb = openpyxl.load_workbook(fpath, data_only=False)
            except Exception as e:
                self._log(f"[警告] 无法加载文件 '{fname}': {e}")
                continue

            if _is_merge_output(wb):
                self._log(f"[跳过] 检测到历史合并产物（含报告/总表页），不作为输入: {fname}")
                continue

            for sname in wb.sheetnames:
                ws = wb[sname]
                unit = SheetUnit(fname, sname, wb, ws)
                try:
                    unit.analyze(self.config)
                    self.sheet_units.append(unit)
                    self._log(describe_region(unit.region, fname, sname))
                except Exception as e:
                    self._log(f"[警告] 跳过异常Sheet分析 [{fname} → {sname}]: {e}")

        non_empty = [u for u in self.sheet_units if not u.is_empty]
        if len(non_empty) < len(self.sheet_units):
            self._log(f"跳过 {len(self.sheet_units) - len(non_empty)} 个空Sheet")
        self.sheet_units = non_empty
        self._col_widths = self._collect_column_widths()

    def run_validation(self) -> None:
        if not self.config.validation_rules:
            self._log("未配置校验规则，跳过数据校验")
            return

        self._log("开始数据校验...")
        self.validation_results = []

        for unit in self.sheet_units:
            try:
                result = validate_sheet(
                    ws=unit.ws,
                    file_name=unit.file_name,
                    sheet_name=unit.sheet_name,
                    body_start=unit.region.body_start,
                    body_end=unit.region.body_end,
                    validation_rules=self.config.validation_rules,
                    strict_mode=self.config.strict_mode,
                )
                self.validation_results.append(result)
                for warn in result.semantic_warnings:
                    self._log(warn)
            except Exception as e:
                self._log(f"[警告] 校验失败 [{unit.file_name} → {unit.sheet_name}]: {e}")
                if self.config.strict_mode:
                    raise

        if not self.config.strict_mode:
            report = validate_all_sheets(self.validation_results)
            self._log(report)
            report_path = os.path.join(os.path.dirname(self.config.output_file), "error_log.txt")
            with open(report_path, "w", encoding="utf-8") as f:
                f.write(report)
            self._log(f"校验报告已保存至: {report_path}")

    def merge(self) -> None:
        if not self.sheet_units:
            self._log("没有可合并的Sheet，退出")
            return

        self._log("开始合并...")
        wb_out = openpyxl.Workbook()
        ws_out = wb_out.active
        ws_out.title = self.config.output_sheet_name

        current_out_row = 1
        total_units = len(self.sheet_units)
        first_unit = self.sheet_units[0]
        header_end = min(self.config.header_rows, first_unit.region.real_max_row)
        # 表头来自首表，按首表表头识别实际天数，日期列宽与填色都以此对齐
        first_attendance_end_col = detect_attendance_day_columns(first_unit.ws, header_end)
        if self.config.copy_column_widths:
            self._apply_column_widths(ws_out, first_attendance_end_col)

        self._log(f"  复制首表表头 [{first_unit.file_name} → {first_unit.sheet_name}] (第1~{header_end}行原样搬运)")
        self._copy_header_block(first_unit, ws_out, header_end)
        current_out_row = header_end + 1

        for idx, unit in enumerate(self.sheet_units):
            try:
                is_first = (idx == 0)
                is_last = (idx == total_units - 1)
                region = unit.region
                max_col = region.real_max_col
                self._log(
                    f"  合并 [{unit.file_name} → {unit.sheet_name}] "
                    f"(第{idx + 1}/{total_units}个, {'首' if is_first else '末' if is_last else '中'})"
                )

                rows_to_copy = self._determine_rows(region, is_first, is_last)
                if is_first:
                    rows_to_copy = range(max(header_end + 1, region.body_start), rows_to_copy.stop)

                unit_out_row_start = current_out_row
                footer_copied = False
                attendance_end_col = detect_attendance_day_columns(unit.ws, unit.region.header_end)

                for src_row in rows_to_copy:
                    if src_row <= header_end and is_first:
                        continue
                    decision = self._classify_body_row(unit, src_row, max_col, is_first, is_last)
                    if decision is not None:
                        self._log(decision)
                        if decision.endswith("[跳过]"):
                            continue
                    if not should_copy_row(unit.ws, src_row, max_col):
                        continue
                    row_offset = current_out_row - src_row
                    copy_row(
                        src_ws=unit.ws,
                        src_row=src_row,
                        target_ws=ws_out,
                        target_row=current_out_row,
                        max_col=max_col,
                        row_offset=row_offset if self.config.adjust_formulas else 0,
                        adjust_formulas=self.config.adjust_formulas,
                        apply_attendance_fill=True,
                        attendance_end_col=attendance_end_col,
                    )
                    if is_last and src_row >= unit.region.footer_start:
                        footer_copied = True
                    current_out_row += 1

                if is_last and not footer_copied and unit.region.footer_start <= unit.region.footer_end:
                    self._copy_footer_block(unit, ws_out, current_out_row)
                    current_out_row += unit.region.footer_end - unit.region.footer_start + 1
                elif rows_to_copy.start < rows_to_copy.stop:
                    src_row_start = rows_to_copy.start
                    src_row_end = rows_to_copy.stop - 1
                    row_offset_for_merge = unit_out_row_start - src_row_start
                    copy_merged_cells(
                        unit.ws, ws_out,
                        src_row_start=src_row_start,
                        src_row_end=src_row_end,
                        target_row_start=unit_out_row_start,
                        row_offset=row_offset_for_merge,
                    )
            except Exception as e:
                self._log(f"[警告] 跳过异常Sheet [{unit.file_name} → {unit.sheet_name}]: {e}")
                continue

        os.makedirs(os.path.dirname(self.config.output_file), exist_ok=True)
        wb_out.save(self.config.output_file)
        self._log(f"合并完成！文件已保存至: {self.config.output_file}")
        self._log(f"输出Sheet: {self.config.output_sheet_name}, 共 {current_out_row - 1} 行")

    def _collect_column_widths(self) -> dict:
        widths = {}
        for unit in self.sheet_units:
            for col in range(1, unit.region.real_max_col + 1):
                letter = openpyxl.utils.get_column_letter(col)
                w = unit.ws.column_dimensions[letter].width
                if w is None:
                    continue
                widths.setdefault(col, []).append(w)
        return widths

    def _apply_column_widths(self, ws_out: openpyxl.worksheet.worksheet.Worksheet, attendance_end_col: int) -> None:
        for col in range(1, max(self._col_widths.keys(), default=0) + 1):
            letter = openpyxl.utils.get_column_letter(col)
            if 1 <= col <= 7:
                # A-G 固定列宽
                ws_out.column_dimensions[letter].width = self.config.col_width_a_g
            elif ATTENDANCE_COL_START <= col <= attendance_end_col:
                # 日期列（H 起，实际天数）固定列宽
                ws_out.column_dimensions[letter].width = self.config.col_width_date
            else:
                vals = self._col_widths.get(col, [])
                if vals:
                    ws_out.column_dimensions[letter].width = sum(vals) / len(vals)

    @staticmethod
    def _determine_rows(region: SheetRegion, is_first: bool, is_last: bool) -> range:
        if is_first and is_last:
            return range(region.header_start, region.footer_end + 1)
        elif is_first:
            return range(region.header_start, region.body_end + 1)
        elif is_last:
            return range(region.body_start, region.footer_end + 1)
        else:
            if region.body_start > region.body_end:
                return range(0, 0)
            return range(region.body_start, region.body_end + 1)

    @staticmethod
    def _copy_header_block(unit: SheetUnit, ws_out: openpyxl.worksheet.worksheet.Worksheet, header_end: int) -> None:
        for src_row in range(1, header_end + 1):
            copy_row(
                src_ws=unit.ws,
                src_row=src_row,
                target_ws=ws_out,
                target_row=src_row,
                max_col=unit.region.real_max_col,
                row_offset=0,
                adjust_formulas=False,
            )
        copy_merged_cells(unit.ws, ws_out, src_row_start=1, src_row_end=header_end, target_row_start=1, row_offset=0)

    @staticmethod
    def _copy_footer_block(unit: SheetUnit, ws_out: openpyxl.worksheet.worksheet.Worksheet, target_start_row: int) -> None:
        footer_rows = list(range(unit.region.footer_start, unit.region.footer_end + 1))
        if not footer_rows:
            return
        src_start = footer_rows[0]
        for offset, src_row in enumerate(footer_rows):
            copy_row(
                src_ws=unit.ws,
                src_row=src_row,
                target_ws=ws_out,
                target_row=target_start_row + offset,
                max_col=unit.region.real_max_col,
                row_offset=0,
                adjust_formulas=False,
            )
        copy_merged_cells(
            unit.ws,
            ws_out,
            src_row_start=src_start,
            src_row_end=footer_rows[-1],
            target_row_start=target_start_row,
            row_offset=target_start_row - src_start,
        )

    @staticmethod
    def _classify_body_row(unit: SheetUnit, src_row: int, max_col: int, is_first: bool, is_last: bool) -> Optional[str]:
        row_text = " ".join(
            str(unit.ws.cell(row=src_row, column=col).value).strip()
            for col in range(1, max_col + 1)
            if unit.ws.cell(row=src_row, column=col).value is not None and str(unit.ws.cell(row=src_row, column=col).value).strip() != ""
        )
        if not row_text:
            return None
        if row_text in ("项目：", "日期"):
            return f"[语义告警] {unit.file_name} → {unit.sheet_name} 第{src_row}行疑似控制行: {row_text} [跳过]"
        if any(kw in row_text for kw in SIGNATURE_KEYWORDS):
            return f"[语义告警] {unit.file_name} → {unit.sheet_name} 第{src_row}行混入表尾签批: {row_text[:120]} [跳过]"
        if any(kw in row_text for kw in LEGEND_KEYWORDS):
            return f"[语义告警] {unit.file_name} → {unit.sheet_name} 第{src_row}行混入表尾图例: {row_text[:120]} [跳过]"
        if not is_effective_row(unit.ws, src_row, max_col):
            return f"[语义告警] {unit.file_name} → {unit.sheet_name} 第{src_row}行只有序号无实质内容 [跳过]"
        if len(row_text) <= 4 and any(ch.isdigit() for ch in row_text):
            return f"[语义告警] {unit.file_name} → {unit.sheet_name} 第{src_row}行疑似只有序号: {row_text} [跳过]"
        return None

    def _log(self, message: str) -> None:
        print(message)
        self._log_lines.append(message)

    def get_log(self) -> str:
        return "\n".join(self._log_lines)

    def execute(self) -> bool:
        try:
            self.config.validate()
            self.load_sheet_units()
            if not self.sheet_units:
                return False
            self.run_validation()
            self.merge()
            return True
        except ValueError as e:
            self._log(f"[错误] {e}")
            return False
        except Exception as e:
            self._log(f"[异常] 合并过程中发生错误: {e}")
            import traceback
            self._log(traceback.format_exc())
            return False
