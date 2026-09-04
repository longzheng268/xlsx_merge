"""
配置模块 - 集中管理所有合并参数与路径配置

职责单一：仅负责配置的定义与校验，不包含任何业务逻辑。
"""

from dataclasses import dataclass
from typing import Dict, Tuple, Type, Optional
import os


@dataclass
class MergeConfig:
    """Excel多表合并的完整配置"""

    # ── 路径配置 ──
    input_dir: str = ""          # 源文件目录
    output_file: str = ""        # 输出文件路径

    # ── 表头/表尾行数 ──
    header_rows: int = 6         # 表头行数（第1~N行为表头，含列头行，第5-6行为合并的列头）
    footer_rows: int = 5         # 表尾行数（不再决定表尾实际大小，改为智能检测）

    # ── 运行模式 ──
    strict_mode: bool = True     # True=严格模式（遇错即停）；False=日志模式（收集全部错误后继续）

    # ── 数据校验规则 ──
    # 键为列索引(1-based)，值为期望类型的元组
    # 例: {2: (int,), 3: str} 表示第2列必须为整数，第3列必须为字符串
    validation_rules: Optional[Dict[int, Tuple[Type, ...]]] = None

    # ── 防乱策略开关 ──
    trim_trailing_empty: bool = True   # 剔除末尾空白行
    adjust_formulas: bool = True     # 公式处理：True=INDIRECT+ROW() 动态绑定；False=相对行号静态平移
    copy_column_widths: bool = True    # 是否复制/计算列宽

    # ── 列宽参数 ──
    col_width_a_g: float = 7.24       # A-G 列固定列宽
    col_width_date: float = 3.5      # 日期列固定列宽（H 起，按实际天数）

    # ── 合并策略 ──
    # "all_in_one" = 所有Sheet合到一张表; "group_by_col_count" = 按列数分组合并
    merge_strategy: str = "all_in_one"

    # ── 输出Sheet名称 ──
    output_sheet_name: str = "Merged_Result"

    def validate(self) -> None:
        """校验配置合法性，不合法时抛出 ValueError"""
        if not self.input_dir:
            raise ValueError("input_dir 不能为空")
        if not self.output_file:
            raise ValueError("output_file 不能为空")
        if not os.path.isdir(self.input_dir):
            raise ValueError(f"input_dir 不存在或不是目录: {self.input_dir}")
        if self.header_rows < 1:
            raise ValueError(f"header_rows 必须 >= 1，当前值: {self.header_rows}")
        if self.footer_rows < 0:
            raise ValueError(f"footer_rows 不能为负，当前值: {self.footer_rows}")
        if self.merge_strategy not in ("all_in_one", "group_by_col_count"):
            raise ValueError(f"不支持的 merge_strategy: {self.merge_strategy}")

    @classmethod
    def from_defaults(cls, project_root: str) -> "MergeConfig":
        """基于项目根目录生成默认配置"""
        return cls(
            input_dir=os.path.join(project_root, "data", "raw_xlsx"),
            output_file=os.path.join(project_root, "data", "out_xlsx", "merged_result.xlsx"),
        )