---
name: excel-csv-format-diagnostics
description: Diagnose format-related read and import failures in Excel (.xlsx) and CSV files. Use when a spreadsheet cannot be opened, imported, or rendered; when a parser reports malformed workbook, XML, comment, encoding, delimiter, or style errors; or when the user needs the exact sheet/cell or CSV row/column causing the failure and an actionable repair.
---

# Excel/CSV 格式诊断

运行 `scripts/diagnose_tabular_format.py <文件路径>`，不要修改源文件。先使用 JSON 结果给出结论；只有用户明确要求时才修复或生成副本。

## 工作流

1. 用脚本检查 `.xlsx`、`.xls` 或 `.csv`；JSON 的 `issues` 是可定位问题，`warnings` 是可能影响兼容性的风险。
2. 若 `location` 有 `sheet` 和 `cell`，按“工作表!单元格”报告；若是 CSV，按“第 N 行，第 M 列”报告。无法映射到单元格的容器/XML 错误要明确标为文件级问题，不能编造坐标。
3. 报告每项的 `code`、`reason`、位置和建议。优先列出 `severity=error`，再列出 warning。
4. 对 Excel，特别关注批注作者为空、线程批注 person 缺显示名、样式索引越界、共享字符串索引越界、无效合并区域和 XML 解析失败。对 CSV，特别关注编码解码失败、NUL 字符、严格 CSV 语法错误和行字段数不一致。
5. 如底层读取器仍失败而诊断没有直接原因，保留原始错误消息、读取器名称与文件级位置；说明这是未归类兼容性错误，并建议使用另一个读取器交叉验证。

## 修复原则

- 先建议 Excel/LibreOffice 中的最小 UI 修复，例如删除并重新创建作者为空的备注。
- 只有在用户授权后，修改 XML 或重写 CSV；始终写入新文件，保留原文件。
- 批注作者为空时，说明会影响严格解析器；修复为非空显示名即可，批注正文无需变动。
- CSV 编码不兼容时，优先建议 UTF-8（含 BOM 可接受）并保留原分隔符、引号和换行风格。
