#!/usr/bin/env python3
"""Audit a construction-organization-design DOCX before Word delivery."""

from __future__ import annotations

import argparse
import collections
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

VAGUE_PATTERNS = {
    "按图施工": r"按(?:施工)?图(?:纸)?(?:要求)?(?:进行)?施工",
    "另详专项方案": r"(?:另详|详见)专项方案",
    "通过专项方案计算确定": r"通过专项方案计算确定",
    "单体另有要求时执行": r"各?单体.*另有要求时.*对应单体执行",
    "原则性编制": r"原则性编制",
    "正式实施参数以": r"正式实施参数以",
}


def qtext(element: ET.Element) -> str:
    return "".join(t.text or "" for t in element.iter(W + "t")).strip()


def normalize(text: str) -> str:
    return re.sub(r"\s+", "", text)


def load_xml(zf: zipfile.ZipFile, name: str) -> ET.Element:
    return ET.fromstring(zf.read(name))


def audit(path: Path) -> dict:
    result = {"file": str(path), "errors": [], "warnings": [], "metrics": {}}
    if not path.is_file():
        result["errors"].append("文件不存在")
        return result
    if path.suffix.lower() != ".docx":
        result["errors"].append("交付文件不是 .docx")
        return result

    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        result["errors"].append("文件不是有效的 DOCX/ZIP 包")
        return result

    names = set(zf.namelist())
    required = {"word/document.xml", "word/styles.xml", "word/settings.xml"}
    missing = sorted(required - names)
    if missing:
        result["errors"].append("缺少核心部件: " + ", ".join(missing))
        return result

    document = load_xml(zf, "word/document.xml")
    styles_root = load_xml(zf, "word/styles.xml")
    settings = load_xml(zf, "word/settings.xml")

    style_names = {}
    for style in styles_root.findall(W + "style"):
        style_id = style.get(W + "styleId", "")
        name_el = style.find(W + "name")
        style_names[style_id] = name_el.get(W + "val", "") if name_el is not None else ""

    tables = document.findall(".//" + W + "tbl")
    paragraphs = document.findall(".//" + W + "p")
    media = [n for n in names if n.startswith("word/media/") and not n.endswith("/")]
    media_bytes = sum(zf.getinfo(n).file_size for n in media)
    drawings = document.findall(".//" + W + "drawing") + document.findall(".//" + W + "pict")
    rels = [n for n in names if n.endswith(".rels")]
    headers = [n for n in names if re.fullmatch(r"word/header\d+\.xml", n)]
    footers = [n for n in names if re.fullmatch(r"word/footer\d+\.xml", n)]

    result["metrics"].update(
        {
            "bytes": path.stat().st_size,
            "tables": len(tables),
            "paragraphs": len(paragraphs),
            "media_files": len(media),
            "media_bytes": media_bytes,
            "drawings": len(drawings),
            "relationships_parts": len(rels),
            "headers": len(headers),
            "footers": len(footers),
        }
    )

    if not tables:
        result["errors"].append("未发现表格")
    if not media and drawings:
        result["errors"].append("文档存在绘图引用但 word/media 为空")

    table_style_counts = collections.Counter()
    body_style_in_table = []
    missing_fixed = []
    missing_header = []
    split_rows = []
    for ti, table in enumerate(tables, 1):
        tbl_pr = table.find(W + "tblPr")
        layout = tbl_pr.find(W + "tblLayout") if tbl_pr is not None else None
        if layout is None or layout.get(W + "type") != "fixed":
            missing_fixed.append(ti)
        rows = table.findall(W + "tr")
        if rows:
            tr_pr = rows[0].find(W + "trPr")
            if tr_pr is None or tr_pr.find(W + "tblHeader") is None:
                missing_header.append(ti)
        for ri, row in enumerate(rows, 1):
            tr_pr = row.find(W + "trPr")
            if tr_pr is None or tr_pr.find(W + "cantSplit") is None:
                split_rows.append(f"{ti}:{ri}")
        for p in table.findall(".//" + W + "p"):
            ppr = p.find(W + "pPr")
            pstyle = ppr.find(W + "pStyle") if ppr is not None else None
            style_id = pstyle.get(W + "val", "") if pstyle is not None else ""
            style_name = style_names.get(style_id, style_id or "(无样式)")
            table_style_counts[style_name] += 1
            if style_name in {"Normal", "正文", "Body Text", "(无样式)"}:
                body_style_in_table.append((ti, qtext(p)[:40], style_name))

    result["metrics"]["table_paragraph_styles"] = dict(table_style_counts)
    if missing_fixed:
        result["errors"].append(f"表格未使用固定布局: {missing_fixed[:20]}")
    if missing_header:
        result["warnings"].append(f"表格首行未设置重复表头: {missing_header[:20]}")
    if split_rows:
        result["warnings"].append(f"表格行允许跨页断行: {split_rows[:20]}")
    if body_style_in_table:
        sample = [f"表{t} {s} {txt}" for t, txt, s in body_style_in_table[:10]]
        result["errors"].append("表格内仍使用正文/空样式: " + " | ".join(sample))

    update_fields = settings.find(W + "updateFields")
    if update_fields is None or update_fields.get(W + "val") not in {"1", "true", "on"}:
        result["warnings"].append("未设置打开 Word 时更新域")
    instr = " ".join((n.text or "") for n in document.findall(".//" + W + "instrText"))
    if "TOC" not in instr:
        result["errors"].append("未发现可更新的 Word 目录域")

    body_text = qtext(document)
    for label, pattern in VAGUE_PATTERNS.items():
        count = len(re.findall(pattern, body_text))
        if count:
            result["warnings"].append(f"疑似模糊甩项表述「{label}」: {count} 处")

    outside_tables = []
    table_paragraph_ids = {id(p) for t in tables for p in t.findall(".//" + W + "p")}
    for p in paragraphs:
        if id(p) not in table_paragraph_ids:
            text = normalize(qtext(p))
            if len(text) >= 12:
                outside_tables.append(text)
    duplicates = [(text, count) for text, count in collections.Counter(outside_tables).items() if count > 1]
    duplicates.sort(key=lambda item: (-item[1], -len(item[0])))
    if duplicates:
        sample = [f"{count}× {text[:60]}" for text, count in duplicates[:15]]
        result["warnings"].append("发现重复正文段落: " + " | ".join(sample))

    zf.close()
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    result = audit(args.docx)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(json.dumps(result["metrics"], ensure_ascii=False, indent=2))
        for item in result["errors"]:
            print("ERROR:", item)
        for item in result["warnings"]:
            print("WARNING:", item)
        if not result["errors"] and not result["warnings"]:
            print("PASS")
    return 2 if result["errors"] else 0


if __name__ == "__main__":
    sys.exit(main())
