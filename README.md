# xlsx_merge

Excel 多表合并工具，基于 `openpyxl` 实现，适合把多个考勤表 xlsx 合并为一张输出表。

## 当前能力

- 读取 `data/raw_xlsx/` 下所有 `.xlsx` 文件
- 遍历每个文件的所有 Sheet
- 首个有效 Sheet 的表头按原样复制到输出
- 保留单元格值、公式、样式、填充、批注、合并单元格
- 合并时保留首表表头、末表表尾
- 对表体做逐行审核，过滤明显无效行、表头/表尾混入行、只有序号无实质内容的行
- 对表体做全量语义校验，发现表头/表尾混入时会写日志
- 支持多个 Sheet 的容错：个别 Sheet 异常时尽量跳过，不影响整份合并
- 输出文件默认写入 `data/out_xlsx/merged_result.xlsx`

## 目录结构

```text
src/
  main.py             # CLI 入口
  merger.py           # 合并主流程
  sheet_analyzer.py   # 表头/表体/表尾分区识别
  cell_utils.py       # 单元格复制、样式、批注、合并单元格
  validator.py        # 数据校验
  config.py           # 配置对象
```

## 运行方式

```bash
python src/main.py
```

## 常用参数

以 `src/main.py` 为准：

- `--input`：输入目录
- `--output`：输出文件路径
- `--header-rows`：表头行数，默认 `6`
- `--footer-rows`：表尾行数，默认 `2`
- `--strict`：严格模式（默认）
- `--no-strict`：日志模式
- `--no-trim`：不剔除末尾空白行
- `--formula-adjust`：启用公式行偏移修正
- `--no-col-width`：不保持列宽
- `--sheet-name`：输出 Sheet 名称

## 配置项（config.py）

`src/config.py` 中的 `MergeConfig` 主要配置如下：

- `input_dir`：输入目录
- `output_file`：输出文件路径
- `header_rows`：表头行数，默认 `6`
- `footer_rows`：表尾行数，默认 `2`
- `strict_mode`：是否严格模式
- `validation_rules`：校验规则 `{列号: 期望类型元组}`
- `trim_trailing_empty`：是否剔除末尾空白行
- `adjust_formulas`：是否修正公式行偏移
- `copy_column_widths`：是否复制/计算列宽
- `merge_strategy`：合并策略，当前支持 `all_in_one`、`group_by_col_count`
- `output_sheet_name`：输出 Sheet 名称

## 合并规则

0. 自动跳过两类非输入文件：
   - 输出文件本身（若输出路径落在输入目录内）；
   - 历史合并产物（首个 Sheet 左上角含“成功合并 / 合并状态 / 合并结果位置”等标记，即上一次合并的输出）。
1. 首个有效 Sheet 的 1~`header_rows` 行原样复制。
2. 中间 Sheet 只合并正文。
3. 最后一个 Sheet 保留表尾。
4. 复制过程中保留：
   - 值 / 公式
   - 样式
   - 批注
   - 合并单元格
   - 行高
5. 表体会做逐行审核和语义校验，遇到明显的表头/表尾/说明行、只有序号无实质内容的行会跳过，并记录日志。

## 列宽

- A-G 列固定为 `7.24`
- 其他列按所有 Sheet 同列列宽取均值

## 注意事项

- 输出文件被 Excel 占用时，保存会失败
- 个别 Sheet 结构不规范时，工具会尽量跳过异常行或异常 Sheet
- 不要把上一次的合并结果（`merged_result.xlsx` 或含“报告/总表”页的文件）放进输入目录，否则其“报告/总表”页会被当成表头复制进来；工具已内置跳过逻辑，但建议输入目录只放原始考勤表
- 这是一个考勤表场景工具，当前规则主要围绕该类表格优化

## 示例

```bash
python src/main.py --input data/raw_xlsx --output data/out_xlsx/merged_result.xlsx
python src/main.py --no-strict --sheet-name Merged_Result
python src/main.py --no-col-width --formula-adjust
```
