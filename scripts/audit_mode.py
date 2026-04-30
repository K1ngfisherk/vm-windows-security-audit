#!/usr/bin/env python3
"""Decide whether an audit request should collect screenshots.

The default is intentionally fast: no screenshots unless the user asks for
visual evidence. Negative screenshot wording wins over positive evidence words.
When screenshots are requested, report image embedding is a separate decision:
screenshots stay in a standalone evidence folder unless the user asks to put
them into the workbook/document.
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path


SCREENSHOT_KEYWORDS = (
    "截图证据",
    "证据截图",
    "截图",
    "截屏",
    "抓屏",
    "图片证据",
    "可视证据",
    "取证",
    "证据",
    "screenshot",
    "screen shot",
    "screen capture",
    "capture evidence",
    "visual evidence",
    "image evidence",
    "evidence",
)

NO_SCREENSHOT_KEYWORDS = (
    "不截图",
    "不用截图",
    "无需截图",
    "不要截图",
    "不需要截图",
    "不截屏",
    "不用证据",
    "无需证据",
    "不要证据",
    "no screenshot",
    "no screenshots",
    "without screenshot",
    "without screenshots",
    "skip screenshot",
    "skip screenshots",
    "no evidence",
)

EMBED_REPORT_KEYWORDS = (
    "截图放到表格",
    "截图放进表格",
    "截图放入表格",
    "截图放在表格",
    "截图写入表格",
    "截图加入表格",
    "截图添加到表格",
    "截图插入表格",
    "截图附到表格",
    "把截图放到表格",
    "把截图放进表格",
    "把截图放入表格",
    "把截图放在表格",
    "把截图写入表格",
    "把截图添加到表格",
    "把截图插入表格",
    "截图放到文档",
    "截图放进文档",
    "截图放入文档",
    "截图放在文档",
    "截图写入文档",
    "截图加入文档",
    "截图添加到文档",
    "截图插入文档",
    "截图附到文档",
    "把截图放到文档",
    "把截图放进文档",
    "把截图放入文档",
    "把截图放在文档",
    "把截图写入文档",
    "把截图添加到文档",
    "把截图插入文档",
    "嵌入截图",
    "插入截图",
    "表格内截图",
    "文档内截图",
    "截图进表",
    "截图进文档",
    "embed screenshot",
    "embed screenshots",
    "insert screenshot",
    "insert screenshots",
    "screenshots in workbook",
    "screenshots into workbook",
    "screenshot in workbook",
    "screenshot into workbook",
    "screenshots in spreadsheet",
    "screenshots into spreadsheet",
    "screenshot in spreadsheet",
    "screenshot into spreadsheet",
    "screenshots in document",
    "screenshots into document",
    "screenshot in document",
    "screenshot into document",
)


NO_EMBED_REPORT_PATTERNS = (
    r"(?:不要|不用|无需|不需要|别|勿|禁止)把?截图.{0,6}(?:表格|文档|报告)",
    r"截图.{0,6}(?:不要|不用|无需|不需要|别|勿|禁止).{0,6}(?:表格|文档|报告)",
    r"(?:不要|不用|无需|不需要|别|勿|禁止).{0,4}(?:嵌入|插入|写入|放入|放进|放到|添加|附到)截图",
    r"(?:do not|don't|dont|without|no)\s+(?:embed|insert|put|attach|include)[^\n.;,]{0,40}screenshot",
    r"screenshot[s]?[^\n.;,]{0,40}(?:not|outside|separate|standalone)[^\n.;,]{0,40}(?:workbook|spreadsheet|document|report)",
)


def no_embed_matches(text: str, embed_keywords: list[str]) -> list[str]:
    matches: list[str] = []
    for pattern in NO_EMBED_REPORT_PATTERNS:
        matches.extend(match.group(0) for match in re.finditer(pattern, text))

    negators = ("不", "不要", "不用", "无需", "不需要", "别", "勿", "禁止", "do not", "don't", "dont", "without", "no ")
    for keyword in embed_keywords:
        folded_keyword = keyword.casefold()
        start = 0
        while True:
            index = text.find(folded_keyword, start)
            if index < 0:
                break
            prefix = text[max(0, index - 32):index]
            if any(negator in prefix for negator in negators):
                matches.append((prefix + folded_keyword).strip())
            start = index + len(folded_keyword)

    deduped: list[str] = []
    seen: set[str] = set()
    for match in matches:
        if match and match not in seen:
            seen.add(match)
            deduped.append(match)
    return deduped


def decide_screenshots(text: str) -> dict[str, object]:
    folded = text.casefold()
    negative = [keyword for keyword in NO_SCREENSHOT_KEYWORDS if keyword.casefold() in folded]
    positive = [keyword for keyword in SCREENSHOT_KEYWORDS if keyword.casefold() in folded]
    embed = [keyword for keyword in EMBED_REPORT_KEYWORDS if keyword.casefold() in folded]
    screenshots = bool(positive) and not negative
    no_embed = no_embed_matches(folded, embed)
    embed_in_report = screenshots and bool(embed) and not no_embed
    if negative:
        reason = "explicit_no_screenshots"
    elif positive:
        reason = "screenshot_keyword"
    else:
        reason = "default_no_screenshots"
    if not screenshots:
        screenshot_mode = "off"
        report_image_mode = "none"
    elif embed_in_report:
        screenshot_mode = "embed_in_report"
        report_image_mode = "embed"
    else:
        screenshot_mode = "separate_evidence"
        report_image_mode = "separate"
    return {
        "screenshots": screenshots,
        "screenshot_mode": screenshot_mode,
        "embed_in_report": embed_in_report,
        "report_image_mode": report_image_mode,
        "report_image_target": "finding_column" if embed_in_report else "",
        "reason": reason,
        "matched_keywords": positive,
        "negative_keywords": negative,
        "embed_keywords": embed,
        "negative_embed_keywords": no_embed,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--text", action="append", default=[], help="Request text to inspect.")
    parser.add_argument("--text-file", type=Path, help="UTF-8 file containing request text.")
    parser.add_argument("--out", type=Path, help="Optional JSON output path.")
    args = parser.parse_args()

    parts = list(args.text)
    if args.text_file:
        parts.append(args.text_file.read_text(encoding="utf-8"))
    result = decide_screenshots("\n".join(parts))
    output = json.dumps(result, ensure_ascii=False, indent=2)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(output, encoding="utf-8")
    print(output)


if __name__ == "__main__":
    main()
