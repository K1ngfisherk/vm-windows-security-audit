#!/usr/bin/env python3
"""Write checklist results into a copied workbook.

Modes:
- text: write finding/status and plain evidence filenames.
- embed-images: write finding/status and insert images in the finding column.

The input plan may be enriched after execution with `finding`, `status`, and
`evidence` fields per item. This script keeps the original workbook untouched.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
from pathlib import Path
from typing import Any

try:
    import openpyxl
    from openpyxl.drawing.image import Image
except ImportError as exc:  # pragma: no cover
    raise SystemExit("openpyxl and Pillow are required for workbook output") from exc


def col_letter(index: int) -> str:
    return openpyxl.utils.get_column_letter(index)


def write_text(cell: Any, text: str) -> None:
    cell.value = text.strip() if text else ""
    cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="top")


def strip_evidence_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\r\n", "\n")
    text = re.sub(r"\n?截图[:：][\s\S]*$", "", text)
    text = re.sub(r"\n?证据[:：][\s\S]*$", "", text)
    text = re.sub(r"\n?row\d{2}_[^\s；;，,]+\.png", "", text, flags=re.I)
    return text.strip()


def report_result(item: dict[str, Any]) -> str:
    for key in ("result", "compliance", "judgement", "judgment"):
        if item.get(key):
            return str(item[key])
    status = str(item.get("status") or "")
    if status and status not in {"captured", "error", "skipped"}:
        return status
    return ""


def fit_image(img: Image, max_width: int, max_height: int) -> Image:
    scale = min(max_width / img.width, max_height / img.height, 1.0)
    img.width = int(img.width * scale)
    img.height = int(img.height * scale)
    return img


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-workbook", required=True, type=Path)
    parser.add_argument("--plan", required=True, type=Path)
    parser.add_argument("--evidence-dir", required=True, type=Path)
    parser.add_argument("--output-workbook", required=True, type=Path)
    parser.add_argument("--mode", choices=["text", "embed-images"], default="text")
    args = parser.parse_args()

    if args.source_workbook.resolve() != args.output_workbook.resolve():
        args.output_workbook.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(args.source_workbook, args.output_workbook)

    plan = json.loads(args.plan.read_text(encoding="utf-8"))
    wb = openpyxl.load_workbook(args.output_workbook)
    ws = wb[plan["sheet"]]
    finding_col = int(plan["columns"].get("finding") or 5)
    result_col = int(plan["columns"].get("result") or 6)

    ws.column_dimensions[col_letter(finding_col)].width = max(ws.column_dimensions[col_letter(finding_col)].width or 20, 72)

    for item in plan["items"]:
        row = int(item["row"])
        finding = item.get("finding") or item.get("observed") or strip_evidence_text(ws.cell(row, finding_col).value)
        status = report_result(item)
        evidence = item.get("evidence") or []

        if args.mode == "text":
            evidence_text = "\n".join(f"截图：{name}" for name in evidence)
            write_text(ws.cell(row, finding_col), "\n".join(part for part in [finding, evidence_text] if part))
            if status:
                write_text(ws.cell(row, result_col), status)
            continue

        write_text(ws.cell(row, finding_col), finding)
        if status:
            write_text(ws.cell(row, result_col), status)

        left_anchor = f"{col_letter(finding_col)}{row}"
        max_image_height = 245 if len(evidence) > 1 else 320
        ws.row_dimensions[row].height = max(ws.row_dimensions[row].height or 20, (max_image_height * 0.75) + 72)

        for index, name in enumerate(evidence):
            image_path = args.evidence_dir / name
            if not image_path.exists():
                continue
            img = fit_image(Image(str(image_path)), 300 if len(evidence) > 1 else 520, max_image_height)
            img.anchor = left_anchor
            ws.add_image(img)
            # openpyxl does not provide simple cell-internal offsets here.
            # For high-fidelity placement, prefer Excel COM or artifact-tool.
            if index > 0:
                break

    wb.save(args.output_workbook)
    print(json.dumps({"output_workbook": str(args.output_workbook), "mode": args.mode}, ensure_ascii=False))


if __name__ == "__main__":
    main()
