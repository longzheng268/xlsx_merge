"""
入口模块 - 解析命令行参数并启动合并流程

用法:
    python main.py [选项]

示例:
    python main.py                          # 使用默认配置
    python main.py --header-rows 5          # 指定表头行数
    python main.py --footer-rows 2          # 指定表尾行数
    python main.py --strict                 # 严格模式（遇错即停）
    python main.py --no-strict              # 日志模式（收集全部错误）
"""

import argparse
import os
import sys

# 将 src 目录加入模块搜索路径，确保可以直接运行 main.py
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import MergeConfig
from merger import MergeEngine


def parse_args() -> argparse.Namespace:
    """解析命令行参数"""
    parser = argparse.ArgumentParser(
        description="Excel多表合并工具 - 合并多个xlsx文件，保留格式、公式与样式",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python main.py                           使用默认配置
  python main.py --header-rows 6           指定表头6行
  python main.py --footer-rows 2           指定表尾2行
  python main.py --no-strict               日志模式（收集全部错误）
  python main.py --no-formula-adjust        强制使用相对行号静态平移
  python main.py --input ./my_data         指定输入目录
  python main.py --output ./result.xlsx    指定输出文件
        """,
    )

    # 路径参数
    parser.add_argument("--input", type=str, default=None,
                        help="输入目录（默认: ../data/raw_xlsx）")
    parser.add_argument("--output", type=str, default=None,
                        help="输出文件路径（默认: ../data/out_xlsx/merged_result.xlsx）")

    # 表头/表尾
    parser.add_argument("--header-rows", type=int, default=6,
                        help="表头行数，含列头行（默认: 6，第5-6行为合并列头）")
    parser.add_argument("--footer-rows", type=int, default=2,
                        help="表尾行数（默认: 2，实际由智能检测决定）")

    # 运行模式
    parser.add_argument("--strict", action="store_true", default=True,
                        help="严格模式：遇到第一个校验错误即停止（默认）")
    parser.add_argument("--no-strict", action="store_true",
                        help="日志模式：收集全部错误后继续，输出 error_log.txt")

    # 防乱策略
    parser.add_argument("--no-trim", action="store_true",
                        help="不剔除末尾空白行（默认会剔除）")
    parser.add_argument("--no-formula-adjust", dest="no_formula_adjust",
                        action="store_true", default=None,
                        help="强制使用相对行号静态平移（不传时由 config.adjust_formulas 决定：True=INDIRECT动态绑定，False=静态平移）")
    parser.add_argument("--no-col-width", action="store_true",
                        help="不保持第一张表的列宽（默认会保持）")

    # 输出Sheet名
    parser.add_argument("--sheet-name", type=str, default="Merged_Result",
                        help="输出Sheet名称（默认: Merged_Result）")

    return parser.parse_args()


def build_config(args: argparse.Namespace) -> MergeConfig:
    """从命令行参数构建配置对象"""
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    config = MergeConfig.from_defaults(project_root)

    # 覆盖路径
    if args.input:
        config.input_dir = os.path.abspath(args.input)
    if args.output:
        config.output_file = os.path.abspath(args.output)

    # 覆盖参数
    config.header_rows = args.header_rows
    config.footer_rows = args.footer_rows
    config.strict_mode = not args.no_strict
    config.trim_trailing_empty = not args.no_trim
    if args.no_formula_adjust is not None:
        config.adjust_formulas = not args.no_formula_adjust
    config.copy_column_widths = not args.no_col_width
    config.output_sheet_name = args.sheet_name

    return config


def main() -> int:
    """主函数，返回退出码"""
    args = parse_args()
    config = build_config(args)

    print("=" * 60)
    print("Excel多表合并工具")
    print("=" * 60)
    print(f"输入目录:   {config.input_dir}")
    print(f"输出文件:   {config.output_file}")
    print(f"表头行数:   {config.header_rows}")
    print(f"表尾行数:   {config.footer_rows}")
    print(f"严格模式:   {'是' if config.strict_mode else '否（日志模式）'}")
    print(f"剔除空行:   {'是' if config.trim_trailing_empty else '否'}")
    print(f"公式修正:   {'是' if config.adjust_formulas else '否'}")
    print(f"保持列宽:   {'是' if config.copy_column_widths else '否'}")
    print("=" * 60)

    engine = MergeEngine(config)
    success = engine.execute()

    if success:
        print("\n[完成] 合并成功！")
        return 0
    else:
        print("\n[失败] 合并失败，请检查日志。")
        return 1


if __name__ == "__main__":
    sys.exit(main())