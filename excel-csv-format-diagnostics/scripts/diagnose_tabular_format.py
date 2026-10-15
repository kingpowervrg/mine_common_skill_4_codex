#!/usr/bin/env python3
"""Read-only structural diagnostics for .xlsx and .csv files."""
from __future__ import annotations
import argparse, csv, io, json, posixpath, re, sys, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

RANGE_REF = re.compile(r"^[A-Z]{1,3}[1-9][0-9]*(?::[A-Z]{1,3}[1-9][0-9]*)?$")
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
def local(tag): return tag.rsplit("}", 1)[-1]
def add(items, severity, code, reason, location=None, suggestion=None):
    item = {"severity": severity, "code": code, "reason": reason}
    if location: item["location"] = {k: v for k, v in location.items() if v is not None}
    if suggestion: item["suggestion"] = suggestion
    items.append(item)
def parse_xml(archive, name, issues):
    try: return ET.fromstring(archive.read(name))
    except (KeyError, ET.ParseError) as exc:
        add(issues, "error", "XLSX_XML_INVALID", f"无法解析 OOXML 部件 {name}: {exc}", {"file": name})
def resolve(base, target): return posixpath.normpath(posixpath.join(posixpath.dirname(base), target)).lstrip("/")
def rels(archive, rel_path, issues):
    root = parse_xml(archive, rel_path, issues)
    return {} if root is None else {node.attrib.get("Id"): node.attrib.get("Target", "") for node in root if local(node.tag) == "Relationship"}
def sheets(archive, issues):
    root = parse_xml(archive, "xl/workbook.xml", issues)
    targets = rels(archive, "xl/_rels/workbook.xml.rels", issues)
    result = {}
    if root is not None:
        for node in root.iter():
            if local(node.tag) == "sheet" and (target := targets.get(node.attrib.get(REL_NS + "id"))):
                result[resolve("xl/workbook.xml", target)] = node.attrib.get("name", "未知工作表")
    return result
def comment_parts(archive, sheet_path, issues):
    rel_path = posixpath.join(posixpath.dirname(sheet_path), "_rels", posixpath.basename(sheet_path) + ".rels")
    if rel_path not in archive.namelist(): return []
    return [resolve(sheet_path, target) for target in rels(archive, rel_path, issues).values() if "comment" in target.lower()]
def diagnose_xlsx(path, issues, warnings):
    try: archive = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError) as exc:
        add(issues, "error", "XLSX_CONTAINER_INVALID", f"不是有效的 .xlsx ZIP 容器: {exc}", {"file": path.name}); return
    with archive:
        bad = archive.testzip()
        if bad: add(issues, "error", "XLSX_ARCHIVE_CRC", f"ZIP 成员校验失败: {bad}", {"file": bad})
        names = set(archive.namelist())
        if "xl/workbook.xml" not in names:
            add(issues, "error", "XLSX_WORKBOOK_MISSING", "缺少 xl/workbook.xml，文件不是完整的 OOXML 工作簿。", {"file": path.name}); return
        style_count = shared_count = 0
        if "xl/styles.xml" in names:
            root = parse_xml(archive, "xl/styles.xml", issues)
            if root is not None:
                style_count = next((sum(1 for child in node if local(child.tag) == "xf") for node in root.iter() if local(node.tag) == "cellXfs"), 0)
        if "xl/sharedStrings.xml" in names:
            root = parse_xml(archive, "xl/sharedStrings.xml", issues)
            if root is not None: shared_count = sum(1 for node in root if local(node.tag) == "si")
        for sheet_path, sheet_name in sheets(archive, issues).items():
            root = parse_xml(archive, sheet_path, issues)
            if root is None: continue
            for node in root.iter():
                if local(node.tag) == "c":
                    location = {"sheet": sheet_name, "cell": node.attrib.get("r"), "file": sheet_path}
                    style = node.attrib.get("s")
                    try:
                        if style is not None and (int(style) < 0 or int(style) >= style_count): add(issues, "error", "XLSX_STYLE_INDEX", f"样式索引 s={style} 超出 cellXfs 范围（0..{max(style_count-1, 0)}）。", location, "复制相邻单元格格式，或重建该单元格样式。")
                    except ValueError: add(issues, "error", "XLSX_STYLE_INDEX", f"样式索引 s={style!r} 不是整数。", location)
                    if node.attrib.get("t") == "s":
                        value = next((child.text for child in node if local(child.tag) == "v"), None)
                        try:
                            if value is None or int(value) < 0 or int(value) >= shared_count: add(issues, "error", "XLSX_SHARED_STRING_INDEX", f"共享字符串索引 {value!r} 超出范围（0..{max(shared_count-1, 0)}）。", location, "重新输入该单元格文本后保存。")
                        except ValueError: add(issues, "error", "XLSX_SHARED_STRING_INDEX", f"共享字符串索引 {value!r} 不是整数。", location)
                elif local(node.tag) == "mergeCell" and not RANGE_REF.fullmatch(node.attrib.get("ref", "")):
                    add(issues, "error", "XLSX_MERGE_RANGE", f"合并区域 {node.attrib.get('ref', '')!r} 不是有效的 A1 范围。", {"sheet": sheet_name, "file": sheet_path}, "取消并重新创建该合并区域。")
            for comment_path in comment_parts(archive, sheet_path, issues):
                comments = parse_xml(archive, comment_path, issues)
                if comments is None: continue
                author_nodes = next((node for node in comments.iter() if local(node.tag) == "authors"), None)
                authors = [] if author_nodes is None else [(child.text or "").strip() for child in author_nodes if local(child.tag) == "author"]
                for node in comments.iter():
                    if local(node.tag) != "comment": continue
                    try: author = authors[int(node.attrib.get("authorId", ""))]
                    except (ValueError, IndexError): author = None
                    if not author: add(issues, "error", "XLSX_COMMENT_AUTHOR_EMPTY", "批注作者为空或 authorId 无效；严格读取器可能要求非空 displayName。", {"sheet": sheet_name, "cell": node.attrib.get("ref"), "file": comment_path}, "删除并重新创建该备注，或将其作者改为非空名称。")
        for person_path in (name for name in names if "persons" in name.lower()):
            root = parse_xml(archive, person_path, issues)
            if root is not None:
                for node in root.iter():
                    if local(node.tag) == "person" and not (node.attrib.get("displayName") or "").strip(): add(issues, "error", "XLSX_THREAD_PERSON_NAME", "线程批注 person 缺少 displayName。", {"file": person_path}, "为该线程批注作者设置非空显示名称。")
def csv_location(text, index): return {"row": text.count("\n", 0, index) + 1, "column": index - text.rfind("\n", 0, index)}
def diagnose_csv(path, issues, warnings):
    raw = path.read_bytes(); text = encoding = None
    for candidate in ("utf-8-sig", "utf-16", "gb18030"):
        try: text, encoding = raw.decode(candidate), candidate; break
        except UnicodeDecodeError: pass
    if text is None:
        try: raw.decode("utf-8")
        except UnicodeDecodeError as exc: add(issues, "error", "CSV_ENCODING", f"无法按 UTF-8、UTF-16 或 GB18030 解码；首个无效字节偏移为 {exc.start}。", {"file": path.name, "row": raw.count(b'\n', 0, exc.start)+1, "column": exc.start-raw.rfind(b'\n', 0, exc.start)}, "将文件转换为 UTF-8 后重试。")
        return
    if encoding != "utf-8-sig": add(warnings, "warning", "CSV_NON_UTF8", f"文件使用 {encoding}；部分读取器仅接受 UTF-8。", {"file": path.name}, "如需跨工具导入，另存为 UTF-8 CSV。")
    if (nul := text.find("\x00")) >= 0: add(issues, "error", "CSV_NUL", "CSV 包含 NUL (U+0000) 字符，常导致解析器拒绝读取。", {"file": path.name, **csv_location(text, nul)}, "删除该字符后另存为 CSV。")
    try: dialect = csv.Sniffer().sniff(text[:8192], delimiters=",;\t|")
    except csv.Error: dialect = csv.excel; add(warnings, "warning", "CSV_DIALECT_GUESS", "无法可靠推断分隔符，按逗号解析。", {"file": path.name}, "如实际使用其他分隔符，请在导入器中显式指定。")
    reader, expected = csv.reader(io.StringIO(text, newline=""), dialect=dialect, strict=True), None
    try:
        for row in reader:
            if expected is None: expected = len(row)
            elif len(row) != expected: add(issues, "error", "CSV_FIELD_COUNT", f"该行有 {len(row)} 列，首行有 {expected} 列。", {"file": path.name, "row": reader.line_num, "column": min(len(row), expected)+1}, "检查缺失或多余的分隔符及未转义的引号。")
    except csv.Error as exc: add(issues, "error", "CSV_SYNTAX", f"CSV 严格解析失败: {exc}", {"file": path.name, "row": reader.line_num or 1}, "检查该行的引号是否成对。")
def main():
    parser = argparse.ArgumentParser(description="Read-only Excel/CSV format diagnostics"); parser.add_argument("file", type=Path); path = parser.parse_args().file
    issues, warnings = [], []
    if not path.is_file(): add(issues, "error", "FILE_NOT_FOUND", "文件不存在或不是常规文件。", {"file": str(path)})
    elif path.suffix.lower() == ".xlsx": diagnose_xlsx(path, issues, warnings)
    elif path.suffix.lower() == ".csv": diagnose_csv(path, issues, warnings)
    else: add(issues, "error", "FORMAT_UNSUPPORTED", "当前脚本仅支持 .xlsx 和 .csv。", {"file": path.name}, "请将旧版 .xls 另存为 .xlsx 后再诊断。")
    print(json.dumps({"file": str(path), "issues": issues, "warnings": warnings, "ok": not issues}, ensure_ascii=False, indent=2)); return int(bool(issues))
if __name__ == "__main__": sys.exit(main())
