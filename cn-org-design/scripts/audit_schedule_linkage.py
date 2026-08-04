#!/usr/bin/env python3
"""Audit key schedule linkage facts in a construction-organization DOCX."""

from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    parser.add_argument("--start", required=True, help="YYYY-MM-DD")
    parser.add_argument("--finish", required=True, help="YYYY-MM-DD")
    parser.add_argument("--peak", required=True, type=int)
    parser.add_argument(
        "--stale-token",
        action="append",
        default=[],
        help="Old date, duration or peak token that must no longer occur; repeatable.",
    )
    return parser.parse_args()


def date_variants(value: str) -> tuple[str, ...]:
    d = dt.date.fromisoformat(value)
    return (
        value,
        value.replace("-", "."),
        f"{d.year}年{d.month}月{d.day}日",
        f"{d.year}年{d.month:02d}月{d.day:02d}日",
    )


def main() -> int:
    args = parse_args()
    if not args.docx.is_file():
        print(f"ERROR: file not found: {args.docx}")
        return 2

    start = dt.date.fromisoformat(args.start)
    finish = dt.date.fromisoformat(args.finish)
    if finish < start:
        print("ERROR: finish precedes start")
        return 2
    inclusive_days = (finish - start).days + 1

    with zipfile.ZipFile(args.docx) as zf:
        names = set(zf.namelist())
        document_xml = zf.read("word/document.xml")
        root = ET.fromstring(document_xml)
        text = "".join(t.text or "" for t in root.findall(".//w:t", NS))
        normalized_text = re.sub(r"\s+", "", text)
        media = [n for n in names if n.startswith("word/media/")]

    errors: list[str] = []
    warnings: list[str] = []

    for label, variants in (("start", date_variants(args.start)), ("finish", date_variants(args.finish))):
        if not any(v in text for v in variants):
            errors.append(f"missing {label} date: {args.start if label == 'start' else args.finish}")

    day_tokens = (f"{inclusive_days}日历天", f"{inclusive_days}天")
    if not any(token in text for token in day_tokens):
        errors.append(f"missing inclusive duration: {inclusive_days} days")

    peak_pattern = re.compile(rf"(?<!\d){args.peak}(?!\d)")
    if not peak_pattern.search(text):
        errors.append(f"missing peak labour value: {args.peak}")

    for heading in ("第八章施工进度计划", "第十三章季节性施工措施", "第十五章施工方案编制计划"):
        if heading not in normalized_text:
            errors.append(f"missing required linked chapter: {heading}")

    for token in args.stale_token:
        if token and token in text:
            errors.append(f"stale token remains: {token}")

    orientations: list[str] = []
    for pg_sz in root.findall(".//w:sectPr/w:pgSz", NS):
        w = int(pg_sz.get(f"{{{W_NS}}}w", "0"))
        h = int(pg_sz.get(f"{{{W_NS}}}h", "0"))
        orientations.append("landscape" if w > h else "portrait")
    landscape_count = orientations.count("landscape")
    if landscape_count == 0:
        warnings.append("no landscape section found for a dense schedule chart")
    if landscape_count > 1:
        warnings.append(f"multiple landscape sections found: {landscape_count}; visually check for duplicate/blank pages")
    if orientations and orientations[-1] != "portrait":
        errors.append("final section is not portrait; later A4 content may remain landscape")
    if not media:
        warnings.append("no embedded media found; confirm the schedule is a native editable chart")

    print(f"DOCX: {args.docx}")
    print(f"Schedule: {args.start} to {args.finish}, inclusive {inclusive_days} days, peak {args.peak}")
    print(f"Sections: {len(orientations)} ({landscape_count} landscape); media: {len(media)}")
    for message in warnings:
        print(f"WARNING: {message}")
    for message in errors:
        print(f"ERROR: {message}")
    if errors:
        return 1
    print("PASS: schedule linkage key facts found")
    return 0


if __name__ == "__main__":
    sys.exit(main())
