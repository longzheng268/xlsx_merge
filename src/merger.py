"""
核心合并引擎模块 - 编排整个合并流程

职责：协调 SheetAnalyzer、Validator、CellUtils 完成合并。
本身不包含具体的单元格操作或校验逻辑。
"""

import os
import sys
import openpyxl
from typing import List, Optional

from config import MergeConfig
from sheet_analyzer import analyze_sheet, describe_region, SheetRegion, SIGNATURE_KEYWORDS, LEGEND_KEYWORDS
from validator import validate_sheet, validate_all_sheets, ValidationResult
from cell_utils import (
    copy_row, copy_column_widths, copy_merged_cells, is_row_empty, should_copy_row
)


# ── 数据结构：单个待合并的Sheet单元 ──

class SheetUnit:
    """一个待合并的Sheet单元（文件+Sheet的组合）"""

    def __init__(self, file_name: str, sheet_name: str,
                 wb: openpyxl.Workbook, ws: openpyxl.worksheet.worksheet.Worksheet):
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.wb = wb
        self.ws = ws
        self.region: Optional[SheetRegion] = None

    def analyze(self, config: MergeConfig) -> None:
        """分析当前Sheet的结构分区"""
        self.region = analyze_sheet(
            self.ws,
            header_rows=config.header_rows,
            footer_rows=config.footer_rows,
            trim_trailing=config.trim_trailing_empty,
        )

    @property
    def is_empty(self) -> bool:
        return self.region is None or self.region.real_max_row == 0


# ── 合并引擎 ──

class MergeEngine:
    """
    Excel多表合并引擎

    工作流：
        1. 扫描输入目录，加载所有xlsx文件
        2. 逐文件逐Sheet分析结构
        3. 可选：执行数据校验
        4. 按序合并到输出工作簿
        5. 保存输出文件
    """

    def __init__(self, config: MergeConfig):
        self.config = config
        self.sheet_units: List[SheetUnit] = []
        self.validation_results: List[ValidationResult] = []
        self._log_lines: List[str] = []
        self._col_widths = self._collect_column_widths()

    # ── 第1步：扫描与加载 ──

    def scan_files(self) -> List[str]:
        """扫描输入目录，返回排序后的xlsx文件列表"""
        input_dir = self.config.input_dir
        files = [
            f for f in os.listdir(input_dir)
            if f.endswith('.xlsx') and not f.startswith('~$')
        ]
        files.sort()
        if not files:
            self._log("未找到可处理的 .xlsx 文件")
        else:
            self._log(f"找到 {len(files)} 个待合并文件:")
            for f in files:
                self._log(f"  - {f}")
        return files

    def load_sheet_units(self) -> None:
        """加载所有文件的所有Sheet为SheetUnit列表"""
        files = self.scan_files()
        self.sheet_units = []

        for fname in files:
            fpath = os.path.join(self.config.input_dir, fname)
            try:
                wb = openpyxl.load_workbook(fpath, data_only=False)
            except Exception as e:
                self._log(f"[警告] 无法加载文件 '{fname}': {e}")
                continue

            for sname in wb.sheetnames:
                ws = wb[sname]
                unit = SheetUnit(fname, sname, wb, ws)
                unit.analyze(self.config)
                self.sheet_units.append(unit)
                self._log(describe_region(unit.region, fname, sname))

        # 过滤掉空Sheet
        non_empty = [u for u in self.sheet_units if not u.is_empty]
        if len(non_empty) < len(self.sheet_units):
            skipped = len(self.sheet_units) - len(non_empty)
            self._log(f"跳过 {skipped} 个空Sheet")
        self.sheet_units = non_empty
        self._col_widths = self._collect_column_widths()

    # ── 第2步：数据校验 ──

    def run_validation(self) -> None:
        """对所有Sheet的数据体执行校验"""
        if not self.config.validation_rules:
            self._log("未配置校验规则，跳过数据校验")
            return

        self._log("开始数据校验...")
        self.validation_results = []

        for unit in self.sheet_units:
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

        # 日志模式下输出汇总报告
        if not self.config.strict_mode:
            report = validate_all_sheets(self.validation_results)
            self._log(report)
            # 将报告写入文件
            report_path = os.path.join(
                os.path.dirname(self.config.output_file), "error_log.txt"
            )
            with open(report_path, "w", encoding="utf-8") as f:
                f.write(report)
            self._log(f"校验报告已保存至: {report_path}")

    # ── 第3步：合并写入 ──

    def merge(self) -> None:
        """执行合并：将所有SheetUnit按序写入输出工作簿"""
        if not self.sheet_units:
            self._log("没有可合并的Sheet，退出")
            return

        self._log("开始合并...")

        wb_out = openpyxl.Workbook()
        ws_out = wb_out.active
        ws_out.title = self.config.output_sheet_name

        current_out_row = 1
        total_units = len(self.sheet_units)

        # 列宽：按所有Sheet的列宽均值，并对常用列强制设置
        if self.config.copy_column_widths:
            self._apply_column_widths(ws_out)

        # 先把首个有效Sheet的表头原样复制到输出，避免重建导致样式/合并/填充丢失
        first_unit = self.sheet_units[0]
        header_end = min(self.config.header_rows, first_unit.region.real_max_row)
        self._log(
            f"  复制首表表头 [{first_unit.file_name} → {first_unit.sheet_name}] "
            f"(第1~{header_end}行原样搬运)"
        )
        self._copy_header_block(first_unit, ws_out, header_end)
        current_out_row = header_end + 1

        # 首表正文与表尾仍按后续逻辑处理，但跳过已复制的表头
        for idx, unit in enumerate(self.sheet_units):
            try:
                is_first = (idx == 0)
                is_last = (idx == total_units - 1)
                region = unit.region
                max_col = region.real_max_col

                self._log(
                    f"  合并 [{unit.file_name} → {unit.sheet_name}] "
                    f"(第{idx + 1}/{total_units}个, "
                    f"{'首' if is_first else '末' if is_last else '中'})"
                )

                rows_to_copy = self._determine_rows(region, is_first, is_last)
                if is_first:
                    rows_to_copy = range(max(header_end + 1, region.body_start), rows_to_copy.stop)

                unit_out_row_start = current_out_row

                for src_row in rows_to_copy:
                    if src_row <= header_end and is_first:
                        continue
                    if not should_copy_row(unit.ws, src_row, max_col):
                        continue
                    row_text = " ".join(
                        str(unit.ws.cell(row=src_row, column=col).value).strip()
                        for col in range(1, max_col + 1)
                        if unit.ws.cell(row=src_row, column=col).value is not None and str(unit.ws.cell(row=src_row, column=col).value).strip() != ""
                    )
                    if not is_first and not is_last:
                        if any(kw in row_text for kw in SIGNATURE_KEYWORDS) or any(kw in row_text for kw in LEGEND_KEYWORDS):
                            continue
                        if row_text == "":
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
                    )
                    current_out_row += 1

                if rows_to_copy.start < rows_to_copy.stop:
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

        # 保存
        os.makedirs(os.path.dirname(self.config.output_file), exist_ok=True)
        wb_out.save(self.config.output_file)
        self._log(f"合并完成！文件已保存至: {self.config.output_file}")
        self._log(f"输出Sheet: {self.config.output_sheet_name}, 共 {current_out_row - 1} 行")

    # ── 辅助方法 ──

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

    def _apply_column_widths(self, ws_out: openpyxl.worksheet.worksheet.Worksheet) -> None:
        for col in range(1, max(self._col_widths.keys(), default=0) + 1):
            letter = openpyxl.utils.get_column_letter(col)
            if 1 <= col <= 7:
                ws_out.column_dimensions[letter].width = 7.24
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
        copy_merged_cells(
            unit.ws,
            ws_out,
            src_row_start=1,
            src_row_end=header_end,
            target_row_start=1,
            row_offset=0,
        )

    def _log(self, message: str) -> None:
        """记录日志并打印"""
        print(message)
        self._log_lines.append(message)

    def get_log(self) -> str:
        """获取完整日志文本"""
        return "\n".join(self._log_lines)

    # ── 完整执行流程 ──

    def execute(self) -> bool:
        """
        执行完整合并流程。

        返回:
            True=成功, False=失败
        """
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