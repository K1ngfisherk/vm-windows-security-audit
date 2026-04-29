#!/usr/bin/env python3
"""Decide whether an audit request should collect screenshots.

The default is intentionally fast: no screenshots unless the user asks for
visual evidence. Negative screenshot wording wins over positive evidence words.
"""

from __future__ import annotations

import argparse
import json
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


def decide_screenshots(text: str) -> dict[str, object]:
    folded = text.casefold()
    negative = [keyword for keyword in NO_SCREENSHOT_KEYWORDS if keyword.casefold() in folded]
    positive = [keyword for keyword in SCREENSHOT_KEYWORDS if keyword.casefold() in folded]
    screenshots = bool(positive) and not negative
    if negative:
        reason = "explicit_no_screenshots"
    elif positive:
        reason = "screenshot_keyword"
    else:
        reason = "default_no_screenshots"
    return {
        "screenshots": screenshots,
        "reason": reason,
        "matched_keywords": positive,
        "negative_keywords": negative,
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
